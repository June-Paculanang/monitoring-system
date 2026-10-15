
import logging

from django.core.mail.backends.smtp import EmailBackend

logger = logging.getLogger(__name__)


class LoggingSMTPEmailBackend(EmailBackend):
    def send_messages(self, email_messages):
        messages = list(email_messages)

        logger.warning(
            "SMTP DIAGNOSTIC: backend called; messages=%s, "
            "host=%s, port=%s, username_configured=%s",
            len(messages),
            self.host,
            self.port,
            bool(self.username),
        )

        try:
            accepted = super().send_messages(messages)

            logger.warning(
                "SMTP DIAGNOSTIC: attempted=%s, accepted=%s",
                len(messages),
                accepted,
            )

            return accepted

        except Exception:
            logger.exception(
                "SMTP DIAGNOSTIC: sending failed."
            )
            raise
