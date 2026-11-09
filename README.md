# Prism frontend — updated My Views integration

This package updates the supplied Prism FastAPI application. It keeps the existing
Databricks identity page, Microsoft Graph employee-profile lookup, Home tab, and
My Views dashboard.

## Project structure

```text
prism-frontend/
├── app.py
├── app.yaml
├── requirements.txt
└── services/
    └── __init__.py
```

## What changed

1. The signed-in Databricks user's email is resolved through Microsoft Graph.
2. The returned Entra `employeeId` is used directly as the BOB `advisor_id`.
3. `ADVISOR_ID_MAP` and its JSON parsing are removed completely.
4. `/api/my-views` calls the separate BOB FastAPI app from the frontend server.
5. The frontend uses `WorkspaceClient()` to authenticate as its Databricks App
   service principal. No manually managed Databricks OAuth token is required.
6. The advisor ID is URL-encoded before it is inserted into the backend path.

## Important configuration

The `GRAPH_CLIENT_SECRET` `valueFrom` must reference the actual secret key/name
configured for this Databricks App. Do not put the client secret in source code.

The Graph app registration needs permission to read the user properties used by
this application, including `employeeId` (typically Microsoft Graph
`User.Read.All` application permission with admin consent).

The Prism frontend app's service principal must have **CAN USE** permission on
the BOB backend app. The backend must accept authenticated Databricks App
requests from that frontend identity.

`CORE_APP_ORIGIN` is retained from your supplied configuration. For a deployed
frontend, set it to the actual permitted browser origin if your application
uses it for CORS or origin validation.

## Identity flow

```text
Authenticated Databricks user
        |
        v
Prism frontend reads trusted proxy email
        |
        v
Microsoft Graph lookup by email / UPN
        |
        v
Entra employeeId
        |
        | employeeId == BOB advisor_id
        v
GET {BOB_BACKEND_URL}/api/v1/advisors/{employeeId}/bob
        |
        v
BOB backend filters and returns that advisor's records
```

The browser calls only `GET /api/my-views`. It does not choose or submit an
advisor ID. The frontend derives it on the server from the authenticated
identity.

## Deploy

1. Update the `GRAPH_CLIENT_SECRET` secret reference in `app.yaml` if the
   Databricks secret is not literally named `secret`.
2. Upload `app.py`, `app.yaml`, `requirements.txt`, and the `services/` folder
   to the Prism frontend Databricks App source.
3. Deploy/restart the Prism frontend app.
4. Confirm the frontend app service principal has **CAN USE** on the BOB backend.
5. Open Prism through its authenticated Databricks Apps URL.
6. Open **My Views**. The frontend calls `/api/my-views`, resolves the Entra
   employee ID, and calls the backend.

## Endpoints

- `GET /` — Prism UI
- `GET /api/user` — Databricks user identity
- `GET /api/auth-debug` — limited identity/token-presence diagnostics
- `GET /api/employee-profile` — Microsoft Graph employee profile
- `GET /api/my-views` — server-side BOB backend integration
- `GET /api/config-check` — non-secret configuration status

## Operational note

The supplied application identifies the interactive user using Databricks Apps
reverse-proxy headers. Keep this app behind the authenticated Databricks Apps
proxy; do not expose it through an unauthenticated alternate ingress.
