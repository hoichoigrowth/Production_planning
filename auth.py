import hashlib
import hmac
import os
import time
import streamlit as st
from typing import Optional

# Rate-limiting state — module-level so it is shared across all concurrent sessions
# (intentional: rate limits must be global, not per-session)
_login_attempts: dict = {}   # email -> [timestamp, ...]
_lockouts: dict = {}         # email -> lockout-expiry timestamp

MAX_ATTEMPTS = 5
LOCKOUT_SECONDS = 15 * 60   # 15 minutes
SESSION_SECONDS = 8 * 3600  # 8 hours


def hash_password(password: str) -> str:
    """Hash a plaintext password with PBKDF2-HMAC-SHA256 and a random salt."""
    salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100_000)
    return salt.hex() + ":" + dk.hex()


def _verify_password(password: str, hashed: str) -> bool:
    """Verify a plaintext password against a PBKDF2-HMAC-SHA256 hash."""
    try:
        salt_hex, dk_hex = hashed.split(":", 1)
        salt = bytes.fromhex(salt_hex)
        dk = bytes.fromhex(dk_hex)
        dk_check = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100_000)
        return hmac.compare_digest(dk, dk_check)
    except Exception:
        return False


def _is_locked_out(email: str) -> tuple:
    """Return (locked: bool, seconds_remaining: float)."""
    expiry = _lockouts.get(email, 0)
    remaining = expiry - time.time()
    if remaining > 0:
        return True, remaining
    # Lockout expired — clear state
    _lockouts.pop(email, None)
    _login_attempts.pop(email, None)
    return False, 0.0


def _record_failed_attempt(email: str) -> int:
    """Record a failed attempt; lock out if threshold reached. Returns attempt count."""
    now = time.time()
    cutoff = now - LOCKOUT_SECONDS
    attempts = [t for t in _login_attempts.get(email, []) if t > cutoff]
    attempts.append(now)
    _login_attempts[email] = attempts
    if len(attempts) >= MAX_ATTEMPTS:
        _lockouts[email] = now + LOCKOUT_SECONDS
    return len(attempts)


def _clear_attempts(email: str) -> None:
    """Clear rate-limit state after a successful login."""
    _login_attempts.pop(email, None)
    _lockouts.pop(email, None)


def _session_expired() -> bool:
    """Return True if the current session is older than SESSION_SECONDS."""
    login_time = st.session_state.get("login_time")
    if login_time is None:
        return True
    return (time.time() - login_time) > SESSION_SECONDS


def authenticate_user() -> bool:
    """
    Show login UI and authenticate the user.

    Returns True when the session is valid and authenticated.
    Returns False and renders the login form otherwise.
    Uses PBKDF2-HMAC-SHA256 password verification, per-email rate limiting
    (5 attempts → 15-minute lockout), and 8-hour session expiry.
    Passwords are stored as hashes in .streamlit/secrets.toml under [users].
    """
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    # Expire sessions older than 8 hours
    if st.session_state.authenticated and _session_expired():
        st.session_state.authenticated = False
        st.session_state.pop("login_time", None)
        st.warning("Your session has expired. Please log in again.")

    if st.session_state.authenticated:
        return True

    # ── Login UI ──────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    .login-header {
        background: linear-gradient(90deg, #ff6b6b, #4ecdc4);
        padding: 2rem;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .login-container {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        border: 1px solid #e0e0e0;
    }
    </style>
    """, unsafe_allow_html=True)

    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown('<div class="login-container">', unsafe_allow_html=True)

            st.subheader("🔐 Employee Access Portal")
            st.write("Please login with your hoichoi corporate email address")

            email = st.text_input(
                "Corporate Email Address",
                placeholder="yourname@hoichoi.tv",
                help="Only @hoichoi.tv email addresses are authorized",
            )
            password = st.text_input(
                "Password",
                type="password",
                help="Enter your corporate password",
            )

            col_a, col_b = st.columns(2)
            with col_a:
                login_clicked = st.button(
                    "🚀 Login", type="primary", use_container_width=True
                )
            with col_b:
                if st.button("ℹ️ Help", use_container_width=True):
                    st.info(
                        "**Need Access?**\n"
                        "- Contact IT department for account setup\n"
                        "- Must use corporate @hoichoi.tv email\n"
                        "- For support: it@hoichoi.tv"
                    )

            if login_clicked:
                if not email or not password:
                    st.error("❌ Please enter both email and password.")
                elif not email.lower().strip().endswith("@hoichoi.tv"):
                    st.error(
                        "❌ Access denied. Only @hoichoi.tv email addresses are authorized."
                    )
                else:
                    locked, remaining = _is_locked_out(email)
                    if locked:
                        mins = int(remaining // 60) + 1
                        st.error(
                            f"❌ Account locked due to too many failed attempts. "
                            f"Try again in {mins} minute(s)."
                        )
                    else:
                        # Look up hash from secrets
                        try:
                            stored_hash = st.secrets.get("users", {}).get(email)
                        except Exception:
                            stored_hash = None

                        if stored_hash and _verify_password(password, stored_hash):
                            _clear_attempts(email)
                            st.session_state.authenticated = True
                            st.session_state.user_email = email
                            st.session_state.user_name = (
                                email.split("@")[0].replace(".", " ").title()
                            )
                            st.session_state.is_admin = email.lower() in [
                                "admin@hoichoi.tv",
                                "sp@hoichoi.tv",
                                "content@hoichoi.tv",
                            ]
                            st.session_state.login_time = time.time()
                            st.success("✅ Login successful! Redirecting...")
                            time.sleep(0.5)
                            st.rerun()
                        else:
                            count = _record_failed_attempt(email)
                            remaining_attempts = max(0, MAX_ATTEMPTS - count)
                            if remaining_attempts > 0:
                                st.error(
                                    f"❌ Invalid credentials. "
                                    f"{remaining_attempts} attempt(s) remaining before lockout."
                                )
                            else:
                                st.error(
                                    f"❌ Account locked for {LOCKOUT_SECONDS // 60} minutes "
                                    f"due to too many failed attempts."
                                )

            st.divider()
            st.markdown(
                """
                <div style='text-align: center; color: #666; font-size: 0.9em;'>
                    <p>🔒 Secure system for hoichoi content review</p>
                    <p>📧 Access restricted to @hoichoi.tv employees only</p>
                    <p>🛡️ All activities are logged for security purposes</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.markdown("</div>", unsafe_allow_html=True)

    return False


def render_user_header() -> None:
    """Render a user info panel and logout button in the sidebar."""
    with st.sidebar:
        st.markdown(
            f"**👤 {st.session_state.get('user_name', 'User')}**"
        )
        st.caption(st.session_state.get("user_email", ""))
        if st.button("Logout", key="auth_logout_btn"):
            for key in ["authenticated", "user_email", "user_name", "is_admin", "login_time"]:
                st.session_state.pop(key, None)
            st.rerun()
