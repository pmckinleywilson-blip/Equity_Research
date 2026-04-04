"""Email delivery service for calendar invites and notifications."""
import base64
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from config import get_settings

logger = logging.getLogger(__name__)
settings = get_settings()


async def send_calendar_invite(
    to_email: str,
    subject: str,
    body_text: str,
    ics_content: bytes,
    unsubscribe_url: str,
) -> bool:
    """Send a calendar invite email with .ics attachment.

    The .ics is sent as both:
    1. A text/calendar MIME part (for auto-population in Outlook/Gmail)
    2. A .ics file attachment (fallback for clients that don't auto-parse)
    """
    if not settings.resend_api_key:
        logger.warning("No Resend API key configured — skipping email to %s", to_email)
        return False

    try:
        import resend
        resend.api_key = settings.resend_api_key

        # Build multipart email
        # The text/calendar part with method=REQUEST triggers auto-add in most clients
        html_body = f"""
        <html><body>
        <p>{body_text}</p>
        <hr>
        <p style="font-size: 12px; color: #666;">
            You're receiving this because you subscribed to event alerts on SP500 Events.<br>
            <a href="{unsubscribe_url}">Unsubscribe</a>
        </p>
        </body></html>
        """

        ics_b64 = base64.b64encode(ics_content).decode("utf-8")

        resend.Emails.send({
            "from": settings.invite_from_email,
            "to": [to_email],
            "subject": subject,
            "html": html_body,
            "headers": {
                "List-Unsubscribe": f"<{unsubscribe_url}>",
                "List-Unsubscribe-Post": "List-Unsubscribe=One-Click",
            },
            "attachments": [
                {
                    "filename": "invite.ics",
                    "content": ics_b64,
                    "content_type": "text/calendar; method=REQUEST",
                }
            ],
        })
        logger.info("Calendar invite sent to %s", to_email)
        return True

    except Exception as e:
        logger.error("Failed to send invite to %s: %s", to_email, e)
        return False
