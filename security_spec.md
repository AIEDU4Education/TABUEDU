# Security Specification: TabuEdu Sections & Students Access Control

## 1. Data Invariants
- **Owner Isolation**: A section can only be read, created, updated, or deleted by its registered owner (`userId == request.auth.uid`). No teacher should be able to view or modify sections belonging to another teacher.
- **Hierarchical Access**: A student document inside `/secciones/{seccionId}/estudiantes/{estudianteId}` can only be read, created, updated, or deleted if the parent section is owned by the current authenticated user.
- **Roster Validation**: Any student record must have a valid identifier `id` and a non-empty name string.
- **Timestamp Integrity**: All sections must capture authentic server-provided timestamps upon creation and update.

## 2. The "Dirty Dozen" Payloads to Block
1. **Creation of Section with Spurred Identity**: Section payload has `userId = "another_teacher_uid"`. Must be blocked.
2. **Access or List Sections to Unauthenticated User**: Read request when guest. Must be blocked.
3. **Injecting shadow fields onto a section**: Section update payload with `{ isApprovedByAdmin: true, description: "hacked" }`. Must be blocked by strict key validation (`keys().hasAll()` & `affectedKeys().hasOnly()`).
4. **Altering section owner**: Section update modifying `userId` once created. Must be blocked (immutability).
5. **Injecting massive strings as section names**: Prevent "Denial of Wallet" resource exhaustion by limiting `nombre.size() <= 100`, `grado.size() <= 50`, `nivel.size() <= 50`.
6. **Injecting invalid document IDs**: Attackers sending document IDs with special characters or excessive sizes. Checked using `isValidId(seccionId)`.
7. **Bypassing Verification status**: Modify section with a tenant that is unverified (if mandatory verified is active; we require `request.auth.token.email_verified == true`).
8. **Orphan student creation**: Enrolling student inside a section that does not exist. Checked using `get(/databases/$(database)/documents/secciones/$(seccionId))` relational verification.
9. **Roster manipulation of another teacher's section**: Creating/deleting a student inside `/secciones/{seccionId}/estudiantes/{estudianteId}` where the parent section does not belong to the active user. Must check `get(/databases/$(database)/documents/secciones/$(seccionId)).data.userId == request.auth.uid`.
10. **Ataining user details without credentials**: Arbitrary list queries for `/secciones` without filtering by `userId == request.auth.uid` to prevent user scraping.
11. **Malicious data types for student roles**: Mutating student roles to unsupported or excessively sized arbitrary fields.
12. **Tampering with creation timestamps**: Forcing client-side future dates instead of `request.time`.

## 3. Test Validation Plan
We'll translate these rules into `firestore.rules`.
Every allow-update must start with the valid model validators.
Email tracking validates `request.auth.token.email_verified == true` if verified.
We will deploy `firestore.rules` and verify the codebase.
