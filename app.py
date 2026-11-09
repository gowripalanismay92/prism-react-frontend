import os
from html import escape

import msal
import requests
from urllib.parse import quote
from databricks.sdk import WorkspaceClient

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse

app = FastAPI(title="Prism Databricks Identity App")


# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------

TENANT_ID = os.environ.get("ENTRA_TENANT_ID")
CLIENT_ID = os.environ.get("GRAPH_CLIENT_ID")
CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET")


# --------------------------------------------------
# DATABRICKS IDENTITY
# --------------------------------------------------

def get_databricks_identity(request: Request) -> dict:
    """
    Read identity headers injected by the Databricks Apps
    reverse proxy.

    Trust these headers only when the app is accessed
    through the authenticated Databricks Apps proxy.
    """

    headers = request.headers

    return {
        "username": headers.get(
            "x-forwarded-preferred-username"
        ),
        "email": headers.get(
            "x-forwarded-email"
        ),
        "user_id": headers.get(
            "x-forwarded-user"
        )
    }


def require_databricks_user(request: Request) -> dict:

    user = get_databricks_identity(request)

    if not user["user_id"]:
        raise HTTPException(
            status_code=401,
            detail=(
                "No Databricks user identity was received. "
                "Open this app through its authenticated "
                "Databricks Apps URL."
            )
        )

    return user


# --------------------------------------------------
# MICROSOFT GRAPH AUTHENTICATION
# --------------------------------------------------

def get_graph_access_token():

    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise HTTPException(
            status_code=500,
            detail="Microsoft Graph configuration is incomplete"
        )

    authority = (
        f"https://login.microsoftonline.com/{TENANT_ID}"
    )

    graph_app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    result = graph_app.acquire_token_for_client(
        scopes=[
            "https://graph.microsoft.com/.default"
        ]
    )

    if "access_token" not in result:
        # Never expose the token or client secret.
        raise HTTPException(
            status_code=502,
            detail="Microsoft Graph authentication failed"
        )

    return result["access_token"]


# --------------------------------------------------
# RETRIEVE EMPLOYEE ID FROM ENTRA ID
# --------------------------------------------------

def get_employee_profile(email: str):

    if not email:
        raise HTTPException(
            status_code=400,
            detail="Advisor email is not available"
        )

    graph_token = get_graph_access_token()

    # Escape apostrophes in the OData string.
    safe_email = email.replace("'", "''")

    graph_url = (
        "https://graph.microsoft.com/v1.0/users"
    )

    params = {
        "$select": (
            "id,displayName,mail,"
            "userPrincipalName,employeeId,"
            "department,jobTitle"
        ),
        "$filter": (
            f"mail eq '{safe_email}' or "
            f"userPrincipalName eq '{safe_email}'"
        )
    }

    response = requests.get(
        graph_url,
        headers={
            "Authorization": f"Bearer {graph_token}"
        },
        params=params,
        timeout=20
    )

    if response.status_code != 200:
        # Do not expose Graph response details or tokens.
        raise HTTPException(
            status_code=502,
            detail="Microsoft Graph user lookup failed"
        )

    users = response.json().get("value", [])

    if len(users) == 0:
        raise HTTPException(
            status_code=404,
            detail="Advisor not found in Entra ID"
        )

    if len(users) > 1:
        raise HTTPException(
            status_code=409,
            detail="Multiple Entra users matched this email"
        )

    user = users[0]

    return {
        "entra_object_id": user.get("id"),
        "display_name": user.get("displayName"),
        "email": (
            user.get("mail")
            or user.get("userPrincipalName")
        ),
        "employee_id": user.get("employeeId"),
        "department": user.get("department"),
        "job_title": user.get("jobTitle")
    }


# --------------------------------------------------
# PRISM HOME PAGE
# --------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):

    user = require_databricks_user(request)

    username = escape(
        user["username"] or "Not provided"
    )

    email = escape(
        user["email"] or "Not provided"
    )

    user_id = escape(
        user["user_id"] or "Not provided"
    )

    return f"""
    <!doctype html>
    <html lang="en">

    <head>
        <meta charset="utf-8">
        <meta name="viewport"
              content="width=device-width, initial-scale=1">

        <title>Prism Advisor Portal</title>

        <style>
            body {{
                font-family: Segoe UI, Arial, sans-serif;
                background: #0f172a;
                color: #e2e8f0;
                min-height: 100vh;
                margin: 0;
                display: grid;
                place-items: center;
            }}

            main {{
                background: #1e293b;
                padding: 28px 36px 36px;
                border-radius: 16px;
                width: min(850px, 92vw);
                box-sizing: border-box;
            }}

            .tab-list {{
                display: flex;
                gap: 8px;
                border-bottom: 1px solid #475569;
                margin: 22px 0 28px;
            }}

            .tab-button {{
                appearance: none;
                border: 0;
                border-bottom: 3px solid transparent;
                background: transparent;
                color: #94a3b8;
                cursor: pointer;
                font: inherit;
                font-weight: 600;
                padding: 12px 18px;
                margin-bottom: -1px;
            }}

            .tab-button:hover {{
                color: #e2e8f0;
            }}

            .tab-button[aria-selected="true"] {{
                color: #f8fafc;
                border-bottom-color: #38bdf8;
            }}

            .tab-panel[hidden] {{
                display: none;
            }}

            .views-empty {{
                margin-top: 18px;
                padding: 28px;
                background: #334155;
                border: 1px solid #475569;
                border-radius: 10px;
                color: #cbd5e1;
            }}

            .views-heading {{ display:flex; justify-content:space-between; align-items:center; gap:16px; }}
            .muted {{ color:#94a3b8; margin-top:4px; }}
            .refresh-button, .bob-search {{ border:1px solid #0e7490; border-radius:8px; background:#0c4a6e; color:#e0f2fe; padding:10px 14px; font:inherit; }}
            .refresh-button {{ cursor:pointer; font-weight:700; }}
            .views-status {{ margin:18px 0; padding:14px 16px; border-radius:9px; background:#164e63; color:#cffafe; }}
            .views-status.error {{ background:#7f1d1d; color:#fee2e2; }}
            .bob-summary-grid {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; margin:18px 0; }}
            .metric-card, .bob-main-panel, .bob-side-panel {{ background:#0c4a6e; border:1px solid #0891b2; border-radius:12px; padding:16px; }}
            .metric-card .metric-label {{ color:#bae6fd; font-size:13px; }}
            .metric-card .metric-value {{ font-size:25px; font-weight:800; margin-top:7px; }}
            .bob-layout {{ display:grid; grid-template-columns:minmax(0,2fr) minmax(230px,.85fr); gap:14px; }}
            .panel-heading {{ display:flex; justify-content:space-between; align-items:center; gap:12px; }}
            .panel-heading h2, .bob-side-panel h2 {{ margin:0 0 6px; font-size:17px; }}
            .bob-search {{ min-width:210px; background:#082f49; }}
            .table-wrap {{ overflow-x:auto; margin-top:14px; }}
            .bob-table {{ width:100%; border-collapse:collapse; font-size:13px; }}
            .bob-table th {{ text-align:left; color:#bae6fd; padding:11px 9px; border-bottom:1px solid #0e7490; white-space:nowrap; }}
            .bob-table td {{ padding:12px 9px; border-bottom:1px solid #155e75; white-space:nowrap; }}
            .status-pill {{ display:inline-block; border-radius:20px; padding:4px 8px; background:#155e75; }}
            .mix-row {{ margin:14px 0; }}
            .mix-label {{ display:flex; justify-content:space-between; gap:8px; font-size:13px; }}
            .mix-track {{ height:8px; background:#164e63; border-radius:10px; margin-top:7px; overflow:hidden; }}
            .mix-fill {{ height:100%; background:#5eead4; border-radius:10px; }}
            .bob-side-panel hr {{ border:0; border-top:1px solid #0e7490; margin:20px 0; }}
            .data-note {{ color:#a5f3fc; font-size:11px; margin-top:24px; }}
            @media (max-width:850px) {{
                .bob-summary-grid {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
                .bob-layout {{ grid-template-columns:1fr; }}
                .panel-heading {{ align-items:flex-start; flex-direction:column; }}
                .bob-search {{ width:100%; box-sizing:border-box; }}
            }}
            .success {{
                color: #34d399;
            }}

            dt {{
                color: #94a3b8;
                margin-top: 14px;
            }}

            dd {{
                margin: 4px 0 0;
                font-weight: 600;
                overflow-wrap: anywhere;
            }}

            .employee {{
                margin-top: 24px;
                padding: 18px;
                background: #334155;
                border-radius: 10px;
            }}
        </style>
    </head>

    <body>

        <main>

            <p>Prism Advisor Portal</p>

            <nav class="tab-list" role="tablist" aria-label="Main navigation">
                <button
                    class="tab-button"
                    id="home-tab"
                    type="button"
                    role="tab"
                    aria-selected="true"
                    aria-controls="home-panel"
                    tabindex="0"
                    onclick="showTab('home')">
                    Home
                </button>
                <button
                    class="tab-button"
                    id="views-tab"
                    type="button"
                    role="tab"
                    aria-selected="false"
                    aria-controls="views-panel"
                    tabindex="-1"
                    onclick="showTab('views')">
                    My Views
                </button>
            </nav>

            <section
                class="tab-panel"
                id="home-panel"
                role="tabpanel"
                aria-labelledby="home-tab">

                <h1>Welcome to Prism</h1>

                <p class="success">
                    Authenticated by Databricks
                </p>

                <h3>Signed-in Databricks user</h3>

                <dl>
                    <dt>Preferred username</dt>
                    <dd>{username}</dd>

                    <dt>Email</dt>
                    <dd>{email}</dd>

                    <dt>Databricks user ID</dt>
                    <dd>{user_id}</dd>
                </dl>

                <div class="employee">
                    <h3>Advisor Employee Profile</h3>
                    <p id="employee-name">Loading advisor profile...</p>
                    <p id="employee-email"></p>
                    <p id="employee-id"></p>
                    <p id="employee-department"></p>
                    <p id="employee-job"></p>
                    <p id="employee-error"></p>
                </div>
            </section>

            <section
                class="tab-panel"
                id="views-panel"
                role="tabpanel"
                aria-labelledby="views-tab"
                hidden>
                <div class="views-heading">
                    <div><h1>My Views</h1><p class="muted">Your advisor book of business</p></div>
                    <button class="refresh-button" type="button" onclick="loadMyViews()">Refresh</button>
                </div>
                <div id="views-status" class="views-status">Select My Views to load your book of business.</div>
                <div id="bob-dashboard" hidden>
                    <div class="bob-summary-grid" id="bob-summary"></div>
                    <div class="bob-layout">
                        <section class="bob-main-panel">
                            <div class="panel-heading">
                                <div><h2>Book of Business</h2><p id="bob-advisor-label" class="muted"></p></div>
                                <input id="bob-search" class="bob-search" type="search" placeholder="Search clients or policies..." oninput="filterBobRecords()">
                            </div>
                            <div class="table-wrap">
                                <table class="bob-table">
                                    <thead><tr><th>Client</th><th>Product</th><th>Status</th><th>Annual premium</th><th>Face amount</th><th>AUM / value</th><th>Next review</th></tr></thead>
                                    <tbody id="bob-records"></tbody>
                                </table>
                            </div>
                        </section>
                        <aside class="bob-side-panel">
                            <h2>Portfolio mix</h2><div id="bob-product-mix"></div>
                            <hr><h2>Policy status</h2><div id="bob-status-mix"></div>
                            <p class="data-note">Synthetic sample BOB data for development and testing.</p>
                        </aside>
                    </div>
                </div>
            </section>

        </main>

        <script>

        function showTab(tabName) {{
            const tabs = ["home", "views"];

            tabs.forEach((name) => {{
                const selected = name === tabName;
                document.getElementById(name + "-tab").setAttribute(
                    "aria-selected", String(selected)
                );
                document.getElementById(name + "-tab").tabIndex = selected ? 0 : -1;
                document.getElementById(name + "-panel").hidden = !selected;
            }});
        }}

        document.querySelectorAll(".tab-button").forEach((tab) => {{
            tab.addEventListener("keydown", (event) => {{
                if (event.key === "ArrowLeft" || event.key === "ArrowRight") {{
                    event.preventDefault();
                    const nextTab = tab.id === "home-tab" ? "views" : "home";
                    document.getElementById(nextTab + "-tab").focus();
                    showTab(nextTab);
                }}
            }});
        }});

        let bobRecords = [];
        function formatCurrency(value) {{
            return new Intl.NumberFormat("en-US", {{style:"currency", currency:"USD", maximumFractionDigits:0}}).format(Number(value || 0));
        }}
        function escapeHtml(value) {{
            return String(value ?? "").replace(/[&<>"']/g, ch => ({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}}[ch]));
        }}
        function setViewsStatus(message, isError = false) {{
            const status = document.getElementById("views-status");
            status.textContent = message;
            status.className = isError ? "views-status error" : "views-status";
            status.hidden = false;
        }}
        function renderMix(elementId, items) {{
            const target = document.getElementById(elementId);
            const total = items.reduce((sum, item) => sum + item.count, 0) || 1;
            target.innerHTML = items.map(item => `
                <div class="mix-row"><div class="mix-label"><span>${{escapeHtml(item.name)}}</span><strong>${{item.count}}</strong></div>
                <div class="mix-track"><div class="mix-fill" style="width:${{(item.count / total) * 100}}%"></div></div></div>
            `).join("");
        }}
        function renderBobRecords(records) {{
            const tbody = document.getElementById("bob-records");
            if (!Array.isArray(records) || records.length === 0) {{
                tbody.innerHTML = '<tr><td colspan="7">No matching records.</td></tr>';
                return;
            }}

            // The BOB API returns snake_case JSON fields.
            tbody.innerHTML = records.map(row => `
                <tr>
                    <td><strong>${{escapeHtml(row.client_name)}}</strong><br><small>${{escapeHtml(row.client_id)}}</small></td>
                    <td>${{escapeHtml(row.product_type)}}</td>
                    <td><span class="status-pill">${{escapeHtml(row.status)}}</span></td>
                    <td>${{formatCurrency(row.annual_premium)}}</td>
                    <td>${{formatCurrency(row.face_amount)}}</td>
                    <td>${{formatCurrency(row.account_value)}}</td>
                    <td>${{escapeHtml(row.next_review_date)}}</td>
                </tr>
            `).join("");
        }}
        function filterBobRecords() {{
            const query = document.getElementById("bob-search").value.toLowerCase();
            renderBobRecords(bobRecords.filter(row => Object.values(row).some(value => String(value).toLowerCase().includes(query))));
        }}
        async function loadMyViews() {{
            setViewsStatus("Loading your advisor book of business...");
            document.getElementById("bob-dashboard").hidden = true;
            try {{
                const response = await fetch("/api/my-views", {{credentials:"same-origin"}});
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || "Unable to load My Views.");
                bobRecords = data.records || [];
                const s = data.summary || {{}};
                const metrics = [
                    ["Total AUM", formatCurrency(s.total_aum)],
                    ["Clients", s.client_count ?? 0],
                    ["In force", s.in_force_count ?? 0],
                    ["Annual premium", formatCurrency(s.total_annual_premium)],
                    ["Face amount", formatCurrency(s.total_face_amount)]
                ];
                document.getElementById("bob-summary").innerHTML = metrics.map(([label,value]) => `<div class="metric-card"><div class="metric-label">${{label}}</div><div class="metric-value">${{value}}</div></div>`).join("");
                document.getElementById("bob-advisor-label").textContent = `Advisor ID: ${{data.advisor_id}} · ${{bobRecords.length}} BOB records`;
                renderBobRecords(bobRecords);
                renderMix("bob-product-mix", data.product_mix || []);
                renderMix("bob-status-mix", data.status_mix || []);
                document.getElementById("bob-dashboard").hidden = false;
                setViewsStatus(`Loaded ${{bobRecords.length}} records for advisor ${{data.advisor_id}}.`);
            }} catch (error) {{ setViewsStatus(error.message, true); }}
        }}
        const initialShowTab = showTab;
        showTab = function(tabName) {{
            initialShowTab(tabName);
            if (tabName === "views" && !window.myViewsRequested) {{
                window.myViewsRequested = true;
                loadMyViews();
            }}
        }};

        async function loadEmployeeProfile() {{

            try {{

                const response = await fetch(
                    "/api/employee-profile"
                );

                const data = await response.json();

                if (!response.ok) {{
                    throw new Error(
                        data.detail || "Profile lookup failed"
                    );
                }}

                document.getElementById(
                    "employee-name"
                ).textContent =
                    "Name: " +
                    (data.display_name || "Not available");

                document.getElementById(
                    "employee-email"
                ).textContent =
                    "Email: " +
                    (data.email || "Not available");

                document.getElementById(
                    "employee-id"
                ).textContent =
                    "Employee ID: " +
                    (data.employee_id || "Not populated in Entra");

                document.getElementById(
                    "employee-department"
                ).textContent =
                    "Department: " +
                    (data.department || "Not available");

                document.getElementById(
                    "employee-job"
                ).textContent =
                    "Job Title: " +
                    (data.job_title || "Not available");

            }} catch (error) {{

                document.getElementById(
                    "employee-name"
                ).textContent =
                    "Unable to load employee profile";

                document.getElementById(
                    "employee-error"
                ).textContent =
                    error.message;

            }}

        }}

        loadEmployeeProfile();

        </script>

    </body>
    </html>
    """


# --------------------------------------------------
# API: DATABRICKS USER
# --------------------------------------------------

@app.get("/api/user")
async def api_user(request: Request):

    user = require_databricks_user(request)

    return {
        "username": user["username"],
        "email": user["email"],
        "user_id": user["user_id"],
        "authenticated": True,
        "identity_provider": "Databricks Apps proxy"
    }


# --------------------------------------------------
# API: AUTH DEBUG
# --------------------------------------------------

@app.get("/api/auth-debug")
async def auth_debug(request: Request):

    user = require_databricks_user(request)

    user_token = request.headers.get(
        "x-forwarded-access-token"
    )

    return {
        "token_received": bool(user_token),
        "token_type": (
            "Bearer" if user_token else None
        ),
        "token_length": (
            len(user_token) if user_token else 0
        ),
        "username": user["username"],
        "email": user["email"],
        "user_id": user["user_id"],
        "authorization_status": (
            "User authorization token received"
            if user_token
            else "User authorization token not received"
        )
    }


# --------------------------------------------------
# API: EMPLOYEE PROFILE
# --------------------------------------------------

@app.get("/api/employee-profile")
def api_employee_profile(request: Request):

    user = require_databricks_user(request)

    email = (
        user["email"]
        or user["username"]
    )

    if not email:
        raise HTTPException(
            status_code=400,
            detail="Advisor email is unavailable"
        )

    profile = get_employee_profile(email)

    return profile


# --------------------------------------------------
# API: MY VIEWS (FRONTEND-TO-BACKEND M2M)
# --------------------------------------------------

def get_advisor_id_for_signed_in_user(request: Request) -> str:
    user = require_databricks_user(request)
    email = user.get("email") or user.get("username")
    profile = get_employee_profile(email)
    employee_id = profile.get("employee_id")
    if not employee_id:
        raise HTTPException(status_code=403, detail="Your Entra profile does not contain an employee ID.")

    # Business rule: Entra employeeId is also the BOB advisor ID.
    # No separate employee-to-advisor translation is needed.
    return str(employee_id).strip()


def get_databricks_m2m_headers() -> dict:
    """
    Authenticate as this Prism frontend Databricks App's service principal.
    Databricks Apps injects the app's OAuth credentials automatically.
    """
    host = os.getenv("DATABRICKS_HOST", "").rstrip("/")
    if not host:
        raise HTTPException(
            status_code=503,
            detail="DATABRICKS_HOST is not configured."
        )

    try:
        workspace_client = WorkspaceClient(host=host)
        return workspace_client.config.authenticate()
    except Exception:
        # Do not expose credential or SDK internals to the browser.
        raise HTTPException(
            status_code=502,
            detail="Unable to obtain Databricks app authentication headers."
        )


@app.get("/api/my-views")
def api_my_views(request: Request):
    advisor_id = get_advisor_id_for_signed_in_user(request)
    backend_url = os.getenv("BOB_BACKEND_URL", "").rstrip("/")
    if not backend_url:
        raise HTTPException(status_code=503, detail="BOB_BACKEND_URL is not configured.")
    headers = get_databricks_m2m_headers()
    safe_advisor_id = quote(advisor_id, safe="")
    try:
        response = requests.get(
            f"{backend_url}/api/v1/advisors/{safe_advisor_id}/bob",
            headers=headers,
            timeout=(5, 30),
        )
    except requests.RequestException:
        raise HTTPException(status_code=502, detail="Unable to reach the BOB backend app.")
    if response.status_code == 404:
        raise HTTPException(status_code=404, detail=f"No BOB records were found for advisor {advisor_id}.")
    if response.status_code in (401, 403):
        raise HTTPException(status_code=502, detail="Backend rejected the frontend app identity. Check CAN USE permission and backend caller restrictions.")
    if response.status_code != 200:
        raise HTTPException(status_code=502, detail="BOB backend returned an error.")
    return response.json()


@app.get("/api/config-check")
async def config_check(request: Request):

    user = require_databricks_user(request)

    return {
        "authenticated": True,
        "tenant_id_configured": bool(
            os.getenv("ENTRA_TENANT_ID")
        ),
        "client_id_configured": bool(
            os.getenv("GRAPH_CLIENT_ID")
        ),
        "client_secret_configured": bool(
            os.getenv("GRAPH_CLIENT_SECRET")
        ),
        "bob_backend_configured": bool(
            os.getenv("BOB_BACKEND_URL")
        ),
        "databricks_host_configured": bool(
            os.getenv("DATABRICKS_HOST")
        )
    }