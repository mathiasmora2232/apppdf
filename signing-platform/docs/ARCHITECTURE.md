# Arquitectura

## Principios

1. La API orquesta; no implementa criptografía.
2. El motor de firma es aislado y reemplazable.
3. Los PDFs originales son inmutables.
4. Cada versión firmada genera un nuevo objeto y hash.
5. Los trabajos pesados se procesan fuera del request HTTP.
6. Coordenadas de campos de firma se guardan normalizadas (0..1).

## Entidades propuestas

- organizations
- users
- certificates
- batches
- documents
- document_versions
- signers
- signature_requests
- signature_fields
- signing_jobs
- audit_events

## Estados base

`DRAFT -> READY -> SIGNING -> PARTIALLY_SIGNED -> COMPLETED`

Errores:

`FAILED`, `CANCELLED`, `CERTIFICATE_EXPIRED`.

## Seguridad pendiente para producción

- Cifrado de secretos y material PKCS#12.
- Nunca persistir contraseñas de certificados.
- TLS entre servicios.
- RBAC por organización.
- Antivirus / content validation al subir PDFs.
- Límites de tamaño y páginas.
- TSA.
- OCSP / CRL.
- Validación completa de cadena X.509.
- Política de retención y borrado.
- Auditoría append-only.
