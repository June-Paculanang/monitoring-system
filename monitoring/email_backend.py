
import json
import logging
import os
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from django.core.mail.backends.base import BaseEmailBackend

logger = logging.getLogger(__name__)


class AppsScriptEmailBackend(BaseEmailBackend):
    def send_messages(self, email_messages):
        if not email_messages:
            return 0

        script_url = os.environ.get("GOOGLE_APPS_SCRIPT_URL")
        relay_secret = os.environ.get("EMAIL_RELAY_SECRET")

        if not script_url or not relay_secret:
            logger.error(
                "Email relay configuration is missing."
            )
            if not self.fail_silently:
                raise RuntimeError(
                    "Email relay environment variables are missing."
                )
            return 0

        sent_count = 0

        for message in email_messages:
            recipients = list(message.recipients())

            if not recipients:
                continue

            payload = {
                "secret": relay_secret,
                "to": ",".join(recipients),
                "subject": message.subject or "",
                "body": message.body or "",
            }

            request = Request(
                script_url,
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                },
                method="POST",
            )

            try:
                with urlopen(request, timeout=25) as response:
                    response_body = response.read().decode("utf-8")

                result = json.loads(response_body)

                if result.get("ok") is not True:
                    raise RuntimeError(
                        "The email relay rejected the request."
                    )

                sent_count += 1
                logger.info(
                    "Password-reset email accepted by the relay."
                )

            except (
                HTTPError,
                URLError,
                TimeoutError,
                OSError,
                ValueError,
                RuntimeError,
            ):
                logger.exception(
                    "Apps Script email delivery failed."
                )

                if not self.fail_silently:
                    raise

        return sent_count
