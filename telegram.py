"""Send a Telegram message. Same bot-token/chat-id pattern as the Girlfriend Day
Notifier project, which reads both from .env.
"""

from __future__ import annotations

import logging
import os

import requests

log = logging.getLogger(__name__)

API = "https://api.telegram.org/bot{token}/sendMessage"
# Telegram rejects messages over 4096 chars; leave room for the chunk counter.
MAX_CHARS = 3900
# Highest TG_BOT_TOKEN_N / TG_CHAT_ID_N pair looked for. Raising it is all that's
# needed to allow another recipient — no other code knows the count.
MAX_ACCOUNTS = 9


def _chunks(lines: list[str]) -> list[str]:
    """Pack lines into messages under the size cap, never splitting a line.

    Keeping lines whole also keeps HTML mode safe: tags never span lines, so a
    chunk boundary can't cut a <b>...</b> pair in half.
    """
    out: list[str] = []
    buf: list[str] = []
    size = 0
    for line in lines:
        if buf and size + len(line) + 1 > MAX_CHARS:
            out.append("\n".join(buf))
            buf, size = [], 0
        buf.append(line)
        size += len(line) + 1
    if buf:
        out.append("\n".join(buf))
    return out


def _accounts() -> list[tuple[str, str]]:
    """Configured (token, chat_id) pairs: the primary plus optional extras.

    Extras are numbered TG_BOT_TOKEN_2 / TG_CHAT_ID_2 upward. The whole range is
    scanned rather than stopping at the first gap, so a pair left blank in .env
    doesn't silently hide every recipient after it.
    """
    token = os.environ.get("TG_BOT_TOKEN")
    chat_id = os.environ.get("TG_CHAT_ID")
    if not (token and chat_id):
        raise RuntimeError("TG_BOT_TOKEN and TG_CHAT_ID must be set in .env")
    accounts = [(token, chat_id)]
    for n in range(2, MAX_ACCOUNTS + 1):
        extra_token = os.environ.get(f"TG_BOT_TOKEN_{n}")
        extra_chat_id = os.environ.get(f"TG_CHAT_ID_{n}")
        if extra_token and extra_chat_id:
            accounts.append((extra_token, extra_chat_id))
    return accounts


def send(text: str, html: bool = False) -> None:
    """Post `text` to every configured chat, splitting if it exceeds Telegram's cap.

    `html=True` sends with Telegram's HTML parse mode (<b>, <code>, ...). Callers
    are responsible for escaping any data in the text with html.escape().
    """
    parts = _chunks(text.split("\n"))
    errors: list[Exception] = []
    for token, chat_id in _accounts():
        try:
            for i, part in enumerate(parts, 1):
                body = part if len(parts) == 1 else f"({i}/{len(parts)})\n{part}"
                payload = {"chat_id": chat_id, "text": body}
                if html:
                    payload["parse_mode"] = "HTML"
                resp = requests.post(API.format(token=token), json=payload, timeout=15)
                resp.raise_for_status()
            log.info("Sent Telegram message to chat %s (%d part(s), %d chars)",
                     chat_id, len(parts), len(text))
        except Exception as e:
            # One account being down must not silence the other, so keep going
            # and re-raise afterwards.
            log.exception("Could not send Telegram message to chat %s", chat_id)
            errors.append(e)
    if errors:
        raise errors[0]
