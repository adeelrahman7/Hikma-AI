# importing main web framework (fastapi), and libraries for handling file uploads,
# sending proper HTTP error responses, and enabling CORS (allowing comms between frontend and backend)
from fastapi import FastAPI, Request, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field, validator, field_validator
from enum import Enum
import re
from typing import Optional, List
import json
import numpy as np
import faiss
from sentence_transformers import SentenceTransformer

# handle requests, file extraction, hashing and timestamps
import requests
import httpx
from starlette.concurrency import run_in_threadpool
import os
import PyPDF2
from pptx import Presentation
import docx
from io import BytesIO
import hashlib
import random
from datetime import datetime
import sqlite3

# adding limits to request size (to prevent overloads)
from slowapi import Limiter
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded
from fastapi.responses import JSONResponse
from fastapi import Header, Depends
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

# adding logging for rate limiting
import logging
from logging.handlers import RotatingFileHandler

# loading environment variables from .env file
from dotenv import load_dotenv
import os
load_dotenv()
API_KEY_ENABLED = os.getenv("API_KEY_ENABLED", "false").lower() == "true"
API_KEY = os.getenv("API_KEY")

# initializing file upload limits
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB
MAX_DOCS = 100  # max number of documents to store in memory

ALLOWED_MIME_TYPES = {
    ".pdf": "application/pdf",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
}

# initializing rate limiter
limiter = Limiter(key_func=get_remote_address)

# setting up logging for rate limiting
os.makedirs("logs", exist_ok=True)
#configuring logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        RotatingFileHandler("logs/api.log", maxBytes=10*1024*1024, backupCount=5),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
logger.info("Starting AI Study Helper API...")

# creating the FastAPI app 
app = FastAPI(
    title = "AI Study Helper",
    description = "An AI-powered chatbot to assist with study materials.",
    version = "1.0.0"
)

# Health check endpoint
# simple endpoint to verify API and Ollama status
@app.get("/health")
def health_check():
    """Check if API and Ollama are running."""
    ollama_status = "unreachable"
    try:
        response = requests.post('http://localhost:11434/api/tags', timeout=2)
        if response.status_code == 200:
            ollama_status = "online"
    except:
        pass
    
    return {
        "status": "healthy",
        "ollama": ollama_status,
        "documents_stored": len(documents_db),
        "max_documents": MAX_DOCS
    }

# adding rate limiting to prevent abuse
app.state.limiter = limiter
@app.exception_handler(RateLimitExceeded)
async def rate_limit_handler(request, exc):
    logger.warning(f"Rate limit exceeded for IP: {request.client.host}, Path: {request.url.path}")
    return JSONResponse(
        status_code=429,
        content={
            "error": "Rate limit exceeded.",
            "detail": "You have sent too many requests. Please try again later.",
            "retry_after": "60 seconds"
        }
    )


# configuring CORS - allowing any frontend to call the api without restrictions
# Note: allow_credentials=False is required alongside allow_origins=["*"];
# browsers reject a wildcard origin combined with credentials=True per spec.
# We authenticate via the X-API-Key header, not cookies, so this is fine.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# mounting static folder
app.mount("/static", StaticFiles(directory="static"), name = "static")
# serve the new frontend's landing page at root (replaces the old demo index.html)
@app.get("/")
def serve_frontend():
    return FileResponse("static/landingpage.html")

@app.get("/favicon.ico", include_in_schema=False)
async def favicon():
    return FileResponse("static/favicon.svg")


# loading the embedder, a place to store the uploaded docs
# and a directory for storing the uploaded files
embedder = None
documents_db = {}
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# --- Persistent storage for uploaded documents ---
# documents_db above is still the in-memory store everything reads from
# during a request — that doesn't change. What's new is that every write to
# it is now mirrored to this SQLite file, and on startup we reload from it
# and rebuild the FAISS index for each document (cheap — it's just
# re-embedding text that's already extracted). This is what makes uploads
# survive a server restart instead of vanishing every time uvicorn reloads.
DB_PATH = "documents.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id TEXT PRIMARY KEY,
            filename TEXT NOT NULL,
            upload_date TEXT NOT NULL,
            text TEXT NOT NULL,
            chunks TEXT NOT NULL,
            topics TEXT NOT NULL,
            page_count INTEGER NOT NULL,
            text_length INTEGER NOT NULL
        )
    """)
    conn.commit()
    conn.close()

def save_document_to_db(doc_id: str, doc: dict):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """INSERT OR REPLACE INTO documents
           (id, filename, upload_date, text, chunks, topics, page_count, text_length)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            doc_id,
            doc["filename"],
            doc["upload_date"],
            doc["text"],
            json.dumps(doc["chunks"]),
            json.dumps(doc["topics"]),
            doc["page_count"],
            doc["text_length"],
        ),
    )
    conn.commit()
    conn.close()

def delete_document_from_db(doc_id: str):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
    conn.commit()
    conn.close()

def load_all_documents_from_db() -> list:
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT id, filename, upload_date, text, chunks, topics, page_count, text_length FROM documents"
    ).fetchall()
    conn.close()
    return rows

# Pydantic models for request bodies (Chat requests, flashcards, questions, study materials, document metadata)
# these define the structure of the data we expect in requests and return in responses

# client MUST send a document_id and material_type

class MaterialType(str, Enum):
    summary = "summary"
    flashcards = "flashcards"
    questions = "questions"
class StudyMaterialRequest(BaseModel):
    document_id: str = Field(..., min_length=32, max_length=32)
    material_type: MaterialType  # e.g., "pdf", "pptx", "docx"
    topic : Optional[str] = Field(None, max_length=100)
    
    @validator("topic")
    def sanitize_topic(cls, v):
        if v:
            return re.sub(r"[^\w\s\-]", "", v)
        return v

class ChatRequest(BaseModel):
    document_id: str = Field(..., min_length=32, max_length=32)
    question: str = Field(..., min_length=3, max_length=500)
    
    @validator("question")
    def sanitize_question(cls, v):
        v = v.replace("```", "")
        v = v.replace("<","").replace(">", "")
        v = v.replace("Ignore previous instructions.", "")
        return v.strip()

class FlashCard(BaseModel):
    front: str
    back: str
    topic: str

class Question(BaseModel):
    question: str
    type: str  # e.g., "multiple_choice", "short_answer"
    options: Optional[List[str]] = None
    answer: str
    explanation: str

# --- Structured schema for /generate/questions (Day 2 rewrite, MCQ-only) ---
# Mirrors study.html's practiceData.mcq shape exactly, so the frontend can
# consume the response with zero field-name translation.
# --- Structured schema for /generate/questions ---
# EXTRACTION comes first: the model's only job is pulling term/definition
# pairs out of the source text (an easy task for a small local model). The
# actual MCQ construction happens in plain Python (see build_mcq_from_facts),
# with zero LLM judgment involved in deciding what's "correct" — this is what
# prevents the model from mixing up related concepts (e.g. git add vs commit)
# and mislabeling the answer key.
class ExtractedFact(BaseModel):
    term: str
    definition: str

    @field_validator("term", "definition")
    @classmethod
    def must_be_real_content(cls, value: str) -> str:
        cleaned = value.strip()
        if len(cleaned) < 3:
            raise ValueError(f"Field '{value}' is too short to be real content")
        return value

class ExtractedFacts(BaseModel):
    facts: List[ExtractedFact] = Field(..., min_length=4)  # need at least 4 to build one MCQ

class MCQItem(BaseModel):
    question: str
    options: List[str] = Field(..., min_length=4, max_length=4)
    correct: int = Field(..., ge=0, le=3)  # 0-based index into options
    source_snippet: str = ""  # the exact extracted text the correct answer came from

    @field_validator("options")
    @classmethod
    def options_must_be_real_answers(cls, options: List[str]) -> List[str]:
        for opt in options:
            cleaned = opt.strip()
            if cleaned.isdigit():
                raise ValueError(f"Option '{opt}' is just a number, not a real answer choice")
            if len(cleaned) < 3:
                raise ValueError(f"Option '{opt}' is too short to be a real answer choice")
        if len({o.strip().lower() for o in options}) < len(options):
            raise ValueError("Duplicate options found in the same question")
        return options

# --- Structured schema for /generate/flashcards ---
# Replaces the old FRONT:/BACK:/--- text parsing (parse_flashcards_by_separator
# etc.) with a validated schema, same reasoning as the MCQ rewrite: guarantees
# well-formed cards instead of hoping the model's formatting holds up.
class FlashcardItem(BaseModel):
    front: str
    back: str

    @field_validator("front", "back")
    @classmethod
    def must_be_real_content(cls, value: str) -> str:
        cleaned = value.strip()
        if len(cleaned) < 3:
            raise ValueError(f"Flashcard field '{value}' is too short to be real content")
        if cleaned.isdigit():
            raise ValueError(f"Flashcard field '{value}' is just a number, not real content")
        return value

class FlashcardSet(BaseModel):
    flashcards: List[FlashcardItem]

# --- Structured schema for /generate/summary ---
# Splits the summary into distinct sections instead of one free-text blob,
# so the frontend can render real headings/lists instead of guessing at
# bullet-point formatting from raw text.
class SummaryContent(BaseModel):
    overview: str
    key_concepts: List[str] = Field(..., min_length=1)
    important_facts: List[str] = Field(..., min_length=1)

# --- Structured schema for the summary's "map" step ---
# Long documents are summarized in two passes (map-reduce): first each chunk
# of the document is condensed into a handful of key points (this schema),
# then all the chunks' key points together are condensed into the final
# SummaryContent above. This lets the whole document be covered, not just
# whatever fits in the first few thousand characters.
class ChunkKeyPoints(BaseModel):
    key_points: List[str] = Field(..., min_length=1, max_length=6)

class StudyMaterial(BaseModel):
    document_id: str
    material_type: str
    content: str
    created_at: str

class DocumentInfo(BaseModel):
    id: str
    filename: str
    upload_date: str
    material_type: str
    page_count: str
    text_length: str
    topics: List[str]

# startup event to load the embedder model and restore any previously
# uploaded documents from disk, so they survive a server restart instead
# of disappearing every time uvicorn reloads.
@app.on_event("startup")
async def startup_event():
    global embedder
    print("Loading embedder model...")
    embedder = SentenceTransformer('all-MiniLM-L6-v2')
    print("Embedder model loaded.")

    init_db()
    rows = load_all_documents_from_db()
    if rows:
        print(f"Restoring {len(rows)} previously uploaded document(s) from disk...")
    for doc_id, filename, upload_date, text, chunks_json, topics_json, page_count, text_length in rows:
        try:
            chunks = json.loads(chunks_json)
            topics = json.loads(topics_json)
            # Re-embedding is the only "recomputation" happening here — the
            # extracted text, chunks and topics are loaded as-is from SQLite;
            # only the FAISS index itself (which can't be stored as plain
            # rows) is rebuilt from those chunks.
            index, embeddings = create_faiss_index(chunks)
            documents_db[doc_id] = {
                "id": doc_id,
                "filename": filename,
                "upload_date": upload_date,
                "text": text,
                "chunks": chunks,
                "index": index,
                "embeddings": embeddings,
                "topics": topics,
                "page_count": page_count,
                "text_length": text_length,
            }
        except Exception as e:
            print(f"Failed to restore document {doc_id} ({filename}): {e}")
    if rows:
        print(f"Restored {len(documents_db)} document(s).")

# helper functions - converting binary files into raw text (critical for embedding and indexing)
# Extract text from PDF files
def extract_text_from_pdf(file_bytes: bytes) -> str:
    pdf_reader = PyPDF2.PdfReader(BytesIO(file_bytes))
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text() + "\n"
    return text

# Extract text from pptx files
def extract_text_from_pptx(file_bytes: bytes) -> str:
    prs = Presentation(BytesIO(file_bytes))
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text

# Extract text from docx files
def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = docx.Document(BytesIO(file_bytes))
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

# Function to chunk text into smaller pieces
# LLMs and embedding models have token limits, so we need to split large texts

# This function splits text into chunks of a specified size (in characters) 
# prevents overly large embedding inputs
def chunk_text(text: str, chunk_size: int = 500) -> List[str]:
    words = text.split()
    chunks = []
    current_chunk = []
    current_length = 0
    for word in words:
        current_chunk.append(word)
        current_length += len(word) + 1  # +1 for the space
        if current_length >= chunk_size:
            chunks.append(" ".join(current_chunk))
            current_chunk = []
            current_length = 0
    if current_chunk:
        chunks.append(" ".join(current_chunk))

    return chunks

# Chunking sized for whole-document LLM processing — NOT the same as
# chunk_text() above, which makes small chunks for embedding search.
# This makes bigger pieces (roughly one "section"/"topic" worth of text)
# so a long document can be walked section-by-section instead of being
# truncated to the first few thousand characters. Each chunk still comfortably
# fits inside the num_ctx window we pass to Ollama for these calls.
def chunk_text_for_llm(text: str, chunk_size_words: int = 700) -> List[str]:
    words = text.split()
    chunks = []
    current = []
    for word in words:
        current.append(word)
        if len(current) >= chunk_size_words:
            chunks.append(" ".join(current))
            current = []
    if current:
        chunks.append(" ".join(current))
    return chunks

# Function to create a FAISS index from text chunks
# this converts texts to embeddings, to float32, normalizes the vectors, and creates the FAISS index and adds vectors to the index.
# this enables semantic similarity search (CORE RAG ARCHITECTURE)
def create_faiss_index(chunks: List[str]):
    embeddings = embedder.encode(chunks, convert_to_tensor=False)
    embeddings = np.array(embeddings).astype("float32")
    faiss.normalize_L2(embeddings)
    dim = embeddings.shape[1]
    index = faiss.IndexFlatIP(dim)
    index.add(embeddings)
    return index, embeddings

# Sends prompt to Ollama API and retrieves AI-generated response, handles timeouts and errors
# Allows us to generate summaries, flashcards, questions, and chat responses
async def generate_with_ollama(prompt: str, model: str = "llama3.2", format_schema: dict = None, options_override: dict = None) -> str:

    payload_options = {
        'temperature': 0.7,
        'num_predict': 1000
    }
    if options_override:
        payload_options.update(options_override)

    payload = {
        'model': model,
        'prompt': prompt,
        'stream': False,
        'options': payload_options
    }
    if format_schema:
        # Constrains Ollama's output to match this JSON schema, instead of
        # just hoping the prompt's formatting instructions get followed.
        payload['format'] = format_schema

    try:
        # Raised from 120s: with full-document chunking, larger num_ctx values,
        # and a "take its time" preference for quality over speed, individual
        # calls on slower CPUs can legitimately take longer than 2 minutes.
        async with httpx.AsyncClient(timeout=300) as client:
            response = await client.post(
                'http://localhost:11434/api/generate',
                json=payload
            )
        if response.status_code == 200:
            return response.json()['response']
        else:
            return "Error: Ollama unavailable."
    except Exception as e:
        return f"Error: {str(e)}"

# Function to extract topics from text using AI
# Analyzes doc, extracts 3-5 main topics, and returns them as a list
# Used for metadata, flashcards, and question generation
async def extract_topics(text: str) -> List[str]:
    # """Extract main topics from text (temporary - no AI)."""
    # # Simple keyword extraction until Ollama is set up
    # keywords = ['biology', 'chemistry', 'physics', 'math', 'history', 
    #             'science', 'anatomy', 'cell', 'molecule', 'equation']
    
    # text_lower = text.lower()
    # found_topics = []
    
    # for keyword in keywords:
    #     if keyword in text_lower:
    #         found_topics.append(keyword.capitalize())
    
    # # Return first 3 found topics, or generic ones
    # if found_topics:
    #     return found_topics[:3]
    # else:
    #     return ["General Study Notes", "Education", "Learning"]


# Uncomment below to use AI-based topic extraction when Ollama is set up

    """Extract main topics from text using AI."""
    prompt = f"""Analyze this text and extract 3-5 main topics/subjects covered.
Return ONLY a comma-separated list of topics, nothing else.

Text:
{text[:2000]}

Topics:"""
    try:
        response = await generate_with_ollama(prompt)
        topics = [t.strip() for t in response.split(',')]
        return topics[:5]
    except Exception as e:
        print(f"Ollama error: {e}")
        return ["General Study Notes", "Education", "Learning"]

def sanitize_for_llm(text:str, max_length:int=3000) -> str:
    """Sanitize text for LLM input by removing special characters and limiting length."""
    text = text.replace("```", "")
    text = re.sub(r"[<>]", "", text)
    text = re.sub(r"Ignore previous instructions.", "", text)
    return text[:max_length].strip()

# --- Deduplication helpers ---
# Nothing about the structured-output/validator pattern stops the model from
# extracting the SAME underlying term or fact twice under slightly different
# wording (e.g. "git add" vs "$ git add"), which would otherwise produce two
# near-identical flashcards or two near-identical practice questions. These
# helpers catch that at the source, right after parsing, before the facts/
# cards are used to build anything else.
def normalize_for_dedup(text: str) -> str:
    """Normalize text so near-identical variants compare as equal.

    Lowercases, strips a leading command-prompt "$", strips punctuation,
    and collapses whitespace — so "Git Add", "git add", and "$ git add"
    all normalize to the same key.
    """
    cleaned = text.strip().lower()
    cleaned = re.sub(r"^\$\s*", "", cleaned)  # strip leading "$ " prompt symbol
    cleaned = re.sub(r"[^a-z0-9\s]", "", cleaned)  # strip punctuation
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned

def dedupe_facts(facts: List["ExtractedFact"]) -> List["ExtractedFact"]:
    """Drop facts whose term or definition normalizes to one already kept.

    Prevents near-duplicate terms from producing two near-identical
    questions, and prevents near-duplicate definitions from being reused
    as confusingly-similar distractors for each other.
    """
    seen_terms = set()
    seen_defs = set()
    deduped = []
    for fact in facts:
        term_key = normalize_for_dedup(fact.term)
        def_key = normalize_for_dedup(fact.definition)
        if term_key in seen_terms or def_key in seen_defs:
            continue
        seen_terms.add(term_key)
        seen_defs.add(def_key)
        deduped.append(fact)
    return deduped

def dedupe_flashcards(cards: List["FlashcardItem"]) -> List["FlashcardItem"]:
    """Drop flashcards whose front or back normalizes to one already kept."""
    seen_fronts = set()
    seen_backs = set()
    deduped = []
    for card in cards:
        front_key = normalize_for_dedup(card.front)
        back_key = normalize_for_dedup(card.back)
        if front_key in seen_fronts or back_key in seen_backs:
            continue
        seen_fronts.add(front_key)
        seen_backs.add(back_key)
        deduped.append(card)
    return deduped

# API key verification dependency (for future use)
def verify_api_key(x_api_key: str = Header(None)):
    """Verify the provided API key."""
    if API_KEY_ENABLED:
        if not x_api_key or x_api_key != API_KEY:
            raise HTTPException(
                status_code=401, 
                detail="Invalid or missing API key."
            )
    return True

# API info endpoint - basic info about the API
# confirms the API is running and provides version and feature list
@app.get("/api/info")
def root():
    return {"message": "Welcome to the AI Study Helper API!",
            "version": "1.0.0",
            "features": [
                "Upload study materials (PDF, PPTX, DOCX)",
                "Generate summary",
                "Create flashcards",
                "Generate practice questions",
                "Chat with your study materials"
            ]
        }

# Endpoint to upload a document
# Validates file type, extracts text, chunks it, creates FAISS index, extracts topics, and stores metadata
# This is the DOCUMENT UPLOAD AND PROCESSING CORE
@app.post("/upload", dependencies=[Depends(verify_api_key)])
@limiter.limit("5/minute") # limit to 5 uploads per minute 
async def upload_document(request: Request, file: UploadFile = File(...)):
    """
    Upload a study material document (PDF, PPTX, DOCX).
    Returns document ID for further processing.
    """
    file_bytes = await file.read()
    logger.info(f"Upload attempt - Filename: {file.filename}, Size: {len(file_bytes)} bytes")
    
    if len (documents_db) >= MAX_DOCS:
        raise HTTPException(
            status_code=429,
            detail="Document storage limit reached. "
        )
        
    #validating file type
    allowed_extensions = ['.pdf', '.pptx', '.docx']
    file_ext = os.path.splitext(file.filename)[1].lower()
    
    if file_ext not in allowed_extensions:
        raise HTTPException(
            status_code=400, detail= f"Unsupported file type. Allowed types: {', '.join(allowed_extensions)}"
            )
    
    # try-catch statements to handle file reading and text extraction
    try: 
        
        # checking file size
        if len(file_bytes) > MAX_FILE_SIZE:
            raise HTTPException(
                status_code=413,
                detail="File too large. Maximum allowed size is 10 MB."
            )
        
        #validate mime type
        if file.content_type and file.content_type != ALLOWED_MIME_TYPES[file_ext]:
            raise HTTPException(
                status_code=400, detail="Invalid MIME type."
            )
        
        # extracting text based on file type
        if file_ext == '.pdf':
            text = extract_text_from_pdf(file_bytes)
        elif file_ext == '.pptx':
            text = extract_text_from_pptx(file_bytes)
        elif file_ext == '.docx':
            text = extract_text_from_docx(file_bytes)
        
        if not text.strip():
            raise HTTPException(
                status_code=400, detail="No extractable text found in the document."
            )
        
        # normalizing and processing text
        text = re.sub(r"\s+", " ", text).strip()
        
        # creating a doc id, chunking text, creating faiss index, extracting topics
        doc_id = hashlib.md5(file_bytes).hexdigest()
        chunks = chunk_text(text, chunk_size=500)
        
        # blocking duplicate uploads
        if doc_id in documents_db:
            raise HTTPException(
                status_code=409, detail="Document already uploaded."
            )
        # empty chunks check
        if not chunks:
            raise HTTPException(
                status_code=400, detail="Document text could not be chunked properly."
            )
        
        # Both do CPU-bound / blocking work, so run them in a worker thread
        # to avoid freezing the event loop during upload
        index, embeddings = await run_in_threadpool(create_faiss_index, chunks)
        topics = await extract_topics(text)
        
        # storing document info in the in-memory db
        documents_db[doc_id] = {
            "id": doc_id,
            "filename": file.filename,
            "upload_date": datetime.utcnow().isoformat(),
            "text": text,
            "chunks": chunks,
            "index": index,
            "embeddings": embeddings,
            "topics": topics,
            "page_count": len(chunks),
            "text_length": len(text),
        }

        # Mirror to SQLite so this document survives a server restart —
        # save_document_to_db() only writes the plain-data fields (not the
        # FAISS index/embeddings, which get rebuilt from chunks on startup).
        save_document_to_db(doc_id, documents_db[doc_id])

        logger.info(f"Document uploaded successfully - ID: {doc_id}, Filename: {file.filename}")

        return {"document_id": doc_id,
                "filename": file.filename,
                "pages_processed": len(chunks),
                "topics_found": topics,
                "message": "Document uploaded and processed successfully."
                }
    except HTTPException:
        # Let intentional HTTP errors (400/409/413/etc.) pass through unchanged
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error processing document: {str(e)}"
        )

# Listing uploaded documents endpoint
# returns metadata for all uploaded documents
@app.get("/documents", dependencies=[Depends(verify_api_key)])
@limiter.limit("30/minute") # limit to 30 document list requests per minute
def list_documents(request: Request):
    """
    List all uploaded documents with their metadata.
    """
    docs = []
    for doc_id, doc in documents_db.items():
        docs.append({
            "id": doc_id,
            "filename": doc["filename"],
            "upload_date": doc["upload_date"],
            "page_count": doc["page_count"],
            "topics": doc["topics"],
        })
    return {"documents": docs, "total": len(docs)}

# Endpoint to generate a summary for a document
# takes document id, retrieves text, constructs prompt, calls LLM, and returns structured summary
@app.post("/generate/summary", dependencies=[Depends(verify_api_key)])
@limiter.limit("10/minute") # limit to 10 summary requests per minute
async def generate_summary(request: Request, req: StudyMaterialRequest):
    """ Generate a structured summary (overview, key concepts, important facts)
    for the specified document, using Ollama's structured outputs feature so
    the response always has real content in a consistent shape. """

    logger.info(f"Summary generation request for Document ID: {req.document_id}")

    if req.document_id not in documents_db:
        logger.warning(f"Document not found for summary generation - ID: {req.document_id}")
        raise HTTPException(status_code=404, detail="Document not found.")

    doc = documents_db[req.document_id]
    # No more truncating to the first 3000 characters — the whole document
    # is walked in sections (map step) and then condensed (reduce step) below,
    # so a long document actually gets summarized in full, not just its start.
    full_text = sanitize_for_llm(doc["text"], max_length=len(doc["text"]) + 1)
    chunks = chunk_text_for_llm(full_text, chunk_size_words=700)
    logger.info(f"Summary: walking {len(chunks)} section(s) of Document ID {req.document_id}")

    MAX_ATTEMPTS = 3

    # --- Map step: condense each section into a few key points ---
    all_key_points: List[str] = []
    for i, chunk in enumerate(chunks, start=1):
        chunk_prompt = f"""List the most important points from this section of a
larger study document, as short factual bullet points. Only use information
found in the text below — do not add outside facts or invent details.

Section:
{chunk}
"""
        for attempt in range(1, MAX_ATTEMPTS + 1):
            response_text = await generate_with_ollama(
                chunk_prompt,
                format_schema=ChunkKeyPoints.model_json_schema(),
                options_override={'temperature': 0.3, 'num_predict': 500, 'num_ctx': 4096}
            )
            if response_text.startswith("Error:"):
                logger.warning(f"Ollama error summarizing section {i}/{len(chunks)} (attempt {attempt}): {response_text}")
                continue
            try:
                chunk_parsed = ChunkKeyPoints.model_validate_json(response_text)
                all_key_points.extend(chunk_parsed.key_points)
                break
            except Exception as e:
                logger.warning(f"Section {i}/{len(chunks)} summary validation failed on attempt {attempt}/{MAX_ATTEMPTS}: {e}")
        # if every attempt for this section failed, we just move on without it
        # rather than failing the whole summary over one bad section

    if not all_key_points:
        logger.error(f"No sections could be summarized for Document ID {req.document_id}")
        raise HTTPException(status_code=502, detail="The AI model is unavailable. Please try again.")

    # --- Reduce step: condense all sections' key points into the final structured summary ---
    combined_points = "\n".join(f"- {p}" for p in all_key_points)
    final_prompt = f"""Below are key points extracted section-by-section from a full
study document. Using ONLY these points, write a study summary. Do not add
outside facts.

Provide:
- overview: a short paragraph (2-4 sentences) covering what the material is about
- key_concepts: a list of the main concepts or terms covered
- important_facts: a list of specific facts, definitions, or details worth remembering

Key points:
{combined_points}
"""

    parsed = None
    last_error = None
    for attempt in range(1, MAX_ATTEMPTS + 1):
        response_text = await generate_with_ollama(
            final_prompt,
            format_schema=SummaryContent.model_json_schema(),
            options_override={'temperature': 0.4, 'num_predict': 900, 'num_ctx': 4096}
        )

        if response_text.startswith("Error:"):
            logger.error(f"Ollama error during final summary reduce (attempt {attempt}): {response_text}")
            raise HTTPException(status_code=502, detail="The AI model is unavailable. Please try again.")

        try:
            parsed = SummaryContent.model_validate_json(response_text)
            break
        except Exception as e:
            last_error = e
            logger.warning(f"Summary validation failed on attempt {attempt}/{MAX_ATTEMPTS}: {e}")

    if parsed is None:
        logger.error(f"All {MAX_ATTEMPTS} attempts failed validation for Document ID {req.document_id}: {last_error}")
        raise HTTPException(status_code=502, detail="Failed to generate a valid summary after multiple attempts. Please try again.")

    logger.info(f"Summary generated for Document ID: {req.document_id} from {len(chunks)} section(s)")

    return {"document_id": req.document_id,
            "material_type": "summary",
            "overview": parsed.overview,
            "key_concepts": parsed.key_concepts,
            "important_facts": parsed.important_facts,
            "created_at": datetime.now().isoformat()
            }

# Endpoint to generate flashcards for a document
# Uses Ollama's structured outputs feature (same pattern as /generate/questions)
# so every card is guaranteed to have real front/back content instead of
# depending on the model following a FRONT:/BACK:/--- text format correctly.
@app.post("/generate/flashcards", dependencies=[Depends(verify_api_key)])
@limiter.limit("10/minute") # limit to 10 flashcard requests per minute
async def generate_flashcards(request: Request, req: StudyMaterialRequest):
    """ Generate flashcards for the specified document."""

    logger.info(f"Flashcard generation request for Document ID: {req.document_id}")

    if req.document_id not in documents_db:
        logger.warning(f"Document not found for flashcard generation - ID: {req.document_id}")
        raise HTTPException(status_code=404, detail="Document not found.")

    doc = documents_db[req.document_id]
    # Walk the whole document in sections instead of truncating to the first
    # 3000 characters, and ask for a handful of cards per section rather than
    # a single fixed count — so a longer/richer document naturally yields
    # more flashcards, roughly scaled to how much content it actually covers.
    full_text = sanitize_for_llm(doc["text"], max_length=len(doc["text"]) + 1)
    # Smaller than the summary's chunk size (700) — a slide deck's extracted
    # text is often sparse, so a big word-count chunk can quietly swallow
    # several distinct topics into one bucket. A smaller chunk here means
    # more, more topic-focused sections, which is what actually drives the
    # total flashcard count up on a bigger/richer document.
    chunks = chunk_text_for_llm(full_text, chunk_size_words=350)
    logger.info(f"Flashcards: walking {len(chunks)} section(s) of Document ID {req.document_id}")

    CARDS_PER_SECTION = 3
    MAX_ATTEMPTS = 3
    all_cards: List[FlashcardItem] = []

    for i, chunk in enumerate(chunks, start=1):
        prompt = f"""Create up to {CARDS_PER_SECTION} flashcards based on the section of
study material below.

Only use information found in the section below — do not invent facts that
aren't supported by it. Each flashcard's "front" should be a clear question
or term, and "back" should be a complete, specific answer or definition of
at least a few words — never a bare number, single word, or placeholder.
If this section doesn't have enough distinct material for {CARDS_PER_SECTION}
good flashcards, return fewer rather than padding with weak ones.

Section:
{chunk}
"""
        for attempt in range(1, MAX_ATTEMPTS + 1):
            response_text = await generate_with_ollama(
                prompt,
                format_schema=FlashcardSet.model_json_schema(),
                # Raised from 500 for the same reason as the questions extraction
                # step — 3 verbose front/back pairs on a dense section can run
                # past a tight token budget and get cut off mid-JSON.
                options_override={'temperature': 0.4, 'num_predict': 900, 'num_ctx': 4096}
            )

            if response_text.startswith("Error:"):
                logger.warning(f"Ollama error generating flashcards for section {i}/{len(chunks)} (attempt {attempt}): {response_text}")
                continue

            try:
                section_parsed = FlashcardSet.model_validate_json(response_text)
                all_cards.extend(section_parsed.flashcards)
                break
            except Exception as e:
                logger.warning(f"Flashcard validation failed for section {i}/{len(chunks)} on attempt {attempt}/{MAX_ATTEMPTS}: {e}")
        # if every attempt for this section failed, just move on without it
        # rather than failing the whole request over one bad section

    deduped_cards = dedupe_flashcards(all_cards)
    if len(deduped_cards) < len(all_cards):
        logger.info(
            f"Removed {len(all_cards) - len(deduped_cards)} duplicate flashcard(s) "
            f"for Document ID: {req.document_id}"
        )

    if len(deduped_cards) < 4:
        logger.error(f"Not enough valid flashcards produced for Document ID {req.document_id}")
        raise HTTPException(status_code=502, detail="Failed to generate enough valid flashcards. Please try again.")

    logger.info(f"Flashcards generated for Document ID: {req.document_id} ({len(deduped_cards)} cards from {len(chunks)} section(s))")

    return {"document_id": req.document_id,
            "material_type": "flashcards",
            "flashcards": [item.model_dump() for item in deduped_cards],
            "count": len(deduped_cards),
            "created_at": datetime.now().isoformat()
            }

# Deterministically builds MCQ items from extracted facts — no LLM judgment
# involved in deciding what's "correct", which is what prevents the model
# from mixing up related concepts (e.g. git add vs commit) and mislabeling
# the answer key. Distractors are other REAL extracted definitions from
# elsewhere in the document, never invented text.
def build_mcq_from_facts(
    facts: List[ExtractedFact],
    num_questions: int = None,
    distractor_pool: List[ExtractedFact] = None,
) -> List[MCQItem]:
    """Build MCQs asking about `facts`, pulling wrong-answer options from
    `distractor_pool` (defaults to `facts` itself). Passing a bigger pool
    (e.g. every fact extracted from the whole document) lets this be called
    per-section while still drawing distractors from the full document,
    instead of only from that one section's handful of facts."""
    pool = distractor_pool if distractor_pool is not None else facts
    if len(pool) < 4:
        raise ValueError("Not enough extracted facts to build multiple-choice questions (need at least 4).")

    chosen = facts if num_questions is None else random.sample(facts, min(num_questions, len(facts)))
    questions = []

    for fact in chosen:
        correct_def = fact.definition
        other_defs = [f.definition for f in pool if f.term != fact.term]
        if len(other_defs) < 3:
            continue  # not enough distinct distractors available for this fact — skip it
        distractors = random.sample(other_defs, 3)

        options = [correct_def] + distractors
        random.shuffle(options)
        correct_index = options.index(correct_def)

        questions.append(MCQItem(
            question=f'What best describes "{fact.term}"?',
            options=options,
            correct=correct_index,
            source_snippet=correct_def
        ))

    return questions

# Endpoint to generate practice questions for a document.
# Uses an extract-then-build approach: the LLM's only job is extracting
# term/definition facts from the source text (an easy task for a small local
# model). The actual questions — including which answer is correct — are
# constructed in plain Python from those facts (see build_mcq_from_facts),
# so answer-key mistakes structurally can't happen the way they could when
# a single LLM call both invents a question AND judges its own answer.
@app.post("/generate/questions", dependencies=[Depends(verify_api_key)])
@limiter.limit("10/minute") # limit to 10 question requests per minute
async def generate_questions(request: Request, req: StudyMaterialRequest):
    """ Generate multiple-choice practice questions for the specified document."""

    logger.info(f"Question generation request for Document ID: {req.document_id}")

    if req.document_id not in documents_db:
        logger.warning(f"Document not found for question generation - ID: {req.document_id}")
        raise HTTPException(status_code=404, detail="Document not found.")

    doc = documents_db[req.document_id]
    # Walk the whole document in sections instead of truncating to the first
    # 3000 characters. Each section is treated as roughly one "topic", and we
    # aim for QUESTIONS_PER_TOPIC questions from each one, so a longer/richer
    # document naturally yields more questions instead of a fixed count.
    # Smaller chunk than the summary's (700) — sparse slide-deck text can
    # otherwise collapse several distinct topics into one bucket, which is
    # what capped the last run at only 3 "topics" for a whole lecture deck.
    full_text = sanitize_for_llm(doc["text"], max_length=len(doc["text"]) + 1)
    chunks = chunk_text_for_llm(full_text, chunk_size_words=350)
    logger.info(f"Questions: walking {len(chunks)} section(s) of Document ID {req.document_id}")

    QUESTIONS_PER_TOPIC = 2
    MAX_TOTAL_QUESTIONS = 20  # cap for big documents — a small doc can still end up with fewer
    MAX_ATTEMPTS = 3

    # --- Extraction pass: pull facts out of each section separately ---
    # section_fact_groups keeps facts grouped by which section they came
    # from (our proxy for "topic"), so we can later pick ~2 per section.
    section_fact_groups: List[List[ExtractedFact]] = []
    for i, chunk in enumerate(chunks, start=1):
        extraction_prompt = f"""Extract the key terms/commands and their definitions from
this section of a larger study document. For each one, give:
- term: the name of the command or concept
- definition: what it does, taken directly from the section below — do not
  paraphrase loosely or add information that isn't there

Extract every distinct term you can find in this section.

Section:
{chunk}
"""
        # A fixed num_predict kept getting outpaced by unusually dense
        # sections (lots of short terms/definitions packed together), which
        # cuts the model off mid-JSON and fails validation every retry — the
        # "EOF while parsing a string" errors. Scale the output budget with
        # how much text is actually in this section instead of guessing one
        # fixed number for every document.
        section_word_count = len(chunk.split())
        extraction_num_predict = min(3000, max(1400, section_word_count * 6))

        section_facts = None
        for attempt in range(1, MAX_ATTEMPTS + 1):
            response_text = await generate_with_ollama(
                extraction_prompt,
                format_schema=ExtractedFacts.model_json_schema(),
                options_override={
                    'temperature': 0.2,   # very low — extraction should be literal, not creative
                    'num_predict': extraction_num_predict,
                    'num_ctx': 6144,
                }
            )

            if response_text.startswith("Error:"):
                logger.warning(f"Ollama error extracting facts for section {i}/{len(chunks)} (attempt {attempt}): {response_text}")
                continue

            try:
                parsed = ExtractedFacts.model_validate_json(response_text)
                section_facts = dedupe_facts(parsed.facts)
                break
            except Exception as e:
                logger.warning(f"Fact extraction failed validation for section {i}/{len(chunks)} on attempt {attempt}/{MAX_ATTEMPTS}: {e}")
        # if every attempt for this section failed, or it had nothing distinct
        # to extract, just move on without it rather than failing the whole request
        if section_facts:
            section_fact_groups.append(section_facts)

    if not section_fact_groups:
        logger.error(f"No sections yielded usable facts for Document ID {req.document_id}")
        raise HTTPException(
            status_code=502,
            detail="Couldn't extract enough distinct facts from this document to build good questions. Try a document with more distinct terms or concepts."
        )

    # Global pool: every unique fact across the whole document, used as the
    # source of distractors so wrong answers stay real and document-grounded
    # even when a section only had one or two facts of its own.
    all_facts = [fact for group in section_fact_groups for fact in group]
    global_facts = dedupe_facts(all_facts)
    if len(all_facts) > len(global_facts):
        logger.info(
            f"Removed {len(all_facts) - len(global_facts)} duplicate fact(s) across sections "
            f"for Document ID: {req.document_id}"
        )

    if len(global_facts) < 4:
        logger.error(f"Not enough unique facts overall for Document ID {req.document_id}")
        raise HTTPException(
            status_code=502,
            detail="Couldn't extract enough distinct facts from this document to build good questions. Try a document with more distinct terms or concepts."
        )

    global_terms = {normalize_for_dedup(f.term) for f in global_facts}

    # --- Build pass: ~2 questions per section/topic, distractors from the whole document ---
    mcq_items: List[MCQItem] = []
    used_terms = set()  # guards against two sections both surfacing the same term
    for group in section_fact_groups:
        # keep only facts from this section that survived the GLOBAL dedup pass
        survivors = [f for f in group if normalize_for_dedup(f.term) in global_terms and normalize_for_dedup(f.term) not in used_terms]
        if not survivors:
            continue
        picked = random.sample(survivors, min(QUESTIONS_PER_TOPIC, len(survivors)))
        for fact in picked:
            used_terms.add(normalize_for_dedup(fact.term))
        try:
            mcq_items.extend(build_mcq_from_facts(picked, distractor_pool=global_facts))
        except ValueError as e:
            logger.warning(f"Skipping a section's questions — {e}")

    if not mcq_items:
        logger.error(f"MCQ construction produced nothing for Document ID {req.document_id}")
        raise HTTPException(status_code=502, detail="Failed to build any valid questions from this document. Please try again.")

    # Cap the total for big/many-section documents — a small document that
    # only produced a handful of questions is left as-is, no padding needed.
    if len(mcq_items) > MAX_TOTAL_QUESTIONS:
        mcq_items = random.sample(mcq_items, MAX_TOTAL_QUESTIONS)

    logger.info(f"Questions generated for Document ID: {req.document_id} ({len(mcq_items)} questions from {len(chunks)} section(s))")
    return {"document_id": req.document_id,
            "material_type": "questions",
            "mcq": [item.model_dump() for item in mcq_items],
            "created_at": datetime.now().isoformat()
            }

# Chat endpoint to interact with uploaded documents (RAG)
# Embedding-based retrieval of relevant chunks, constructs chat prompt, calls LLM, and returns response
@app.post("/chat", dependencies=[Depends(verify_api_key)])
@limiter.limit("20/minute") # limit to 20 chat requests per minute
async def chat_with_document(request: Request, req: ChatRequest):
    """ Ask questions about the uploaded document."""
    logger.info(f"Chat request for Document ID: {req.document_id}, Question: {req.question[:50]}")
    
    if req.document_id not in documents_db:
        logger.warning(f"Document not found for chat - ID: {req.document_id}")
        raise HTTPException(status_code=404, detail="Document not found.")
    
    doc = documents_db[req.document_id]
    
    # search for relevant chunks using FAISS
    # embedder.encode is blocking, so run it off the event loop
    query_vector = await run_in_threadpool(embedder.encode, [req.question], convert_to_tensor=False)
    query_vector = np.array(query_vector).astype("float32")
    faiss.normalize_L2(query_vector)
    
    scores, indices = doc["index"].search(query_vector, k=3)  # retrieve top 3 relevant chunks
    
    # get relevant text chunks/context
    raw_context = "\n\n".join([doc["chunks"][idx] for idx in indices[0]])
    context = sanitize_for_llm(raw_context, max_length=1500)  # limit context size
    
    # generate an answer using the context
    prompt = f"""Use the following context from the study material to answer the question.
Context:
{context}

Question: {req.question}
Provide a clear, educational answer. If the context does not contain the answer, say so.

Answer:"""
    answer = await generate_with_ollama(prompt)
    
    logger.info(f"Chat response generated for Document ID: {req.document_id}, Confidence Score: {float(scores[0][0])}")
    return {"document_id": req.document_id,
            "question": req.question,
            "answer": answer,
            "confidence_score": float(scores[0][0]),
            "sources_used": len(indices[0])
            }

# Endpoint to delete an uploaded document
# removes document and its data from in-memory db
@app.delete("/documents/{document_id}", dependencies=[Depends(verify_api_key)])
@limiter.limit("10/minute") # limit to 10 delete requests per minute
def delete_document(request: Request, document_id: str):
    """ Delete an uploaded document and its data."""
    logger.info(f"Delete request for Document ID: {document_id}")
    
    if document_id not in documents_db:
        logger.warning(f"Document not found for deletion - ID: {document_id}")
        raise HTTPException(status_code=404, detail="Document not found.")
    
    filename = documents_db[document_id]["filename"]
    del documents_db[document_id]
    delete_document_from_db(document_id)

    logger.info(f"Document deleted successfully - ID: {document_id}, Filename: {filename}")
    return {"message": f"Document '{filename}' and its data have been deleted."}

# Catch-all mount for the rest of the frontend's static files (upload.html,
# study.html, css/styles.css, assets/...). Registered LAST, after every API
# route above, so it only serves paths that don't match an actual endpoint —
# it never shadows /upload, /generate/*, /chat, etc. This is what makes
# relative links like href="upload.html" and href="css/styles.css" resolve
# correctly when the page is loaded from "/" instead of needing "/static/..."
# in every href.
app.mount("/", StaticFiles(directory="static"), name="static-root")

# To run the app, use the command:
# uvicorn api:app --host
# This will start the FastAPI server on the specified host and port.
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)