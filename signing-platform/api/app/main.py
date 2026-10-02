from enum import Enum
from uuid import UUID, uuid4

from fastapi import FastAPI
from pydantic import BaseModel, Field

app = FastAPI(title="Document Signing API", version="0.1.0")


class WorkflowMode(str, Enum):
    sequential = "SEQUENTIAL"
    parallel = "PARALLEL"
    mixed = "MIXED"


class SignatureField(BaseModel):
    signer_id: UUID
    page: int = Field(ge=1)
    x: float = Field(ge=0, le=1)
    y: float = Field(ge=0, le=1)
    width: float = Field(gt=0, le=1)
    height: float = Field(gt=0, le=1)


class CreateBatchRequest(BaseModel):
    name: str = Field(min_length=1, max_length=160)
    workflow_mode: WorkflowMode = WorkflowMode.sequential


class BatchResponse(BaseModel):
    id: UUID
    name: str
    workflow_mode: WorkflowMode
    status: str


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/v1/batches", response_model=BatchResponse, status_code=201)
def create_batch(payload: CreateBatchRequest):
    # TODO: persistir en PostgreSQL.
    return BatchResponse(
        id=uuid4(),
        name=payload.name,
        workflow_mode=payload.workflow_mode,
        status="DRAFT",
    )


@app.post("/api/v1/documents/{document_id}/signature-fields")
def add_signature_field(document_id: UUID, field: SignatureField):
    # TODO: persistir y validar dimensiones reales del PDF.
    return {"document_id": document_id, "field": field, "status": "created"}
