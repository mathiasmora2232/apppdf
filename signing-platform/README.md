# Document Signing Platform

Base de una plataforma de firma documental por lotes, multifirmante y multipágina.

## Arquitectura

- `web`: Next.js + TypeScript
- `api`: FastAPI
- `signature-engine`: Java 21 + Apache PDFBox + Bouncy Castle
- PostgreSQL: metadatos y workflow
- Redis: cola/estado de trabajos
- MinIO: almacenamiento S3-compatible

## Flujo MVP

1. Crear lote.
2. Subir uno o más PDF.
3. Registrar firmantes.
4. Definir campos de firma por página usando coordenadas normalizadas.
5. Encolar trabajo.
6. El motor Java aplica la firma criptográfica.
7. Guardar resultado y auditoría.

## Arranque local

```bash
cp .env.example .env
docker compose up --build
```

Servicios:

- Web: http://localhost:3000
- API: http://localhost:8000/docs
- Signature Engine: http://localhost:8080
- MinIO: http://localhost:9001

> Base inicial. No usar en producción sin implementar gestión segura de claves, validación X.509, TSA/OCSP/CRL y hardening.
