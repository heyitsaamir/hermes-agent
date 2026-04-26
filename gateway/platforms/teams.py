"""
Microsoft Teams platform adapter.

Uses the microsoft-teams-apps SDK for authentication and activity processing.
Runs an aiohttp webhook server to receive messages from Teams.
Proactive messaging (send, typing) uses the SDK's App.send() method.

Requires:
    pip install microsoft-teams-apps aiohttp
    TEAMS_CLIENT_ID, TEAMS_CLIENT_SECRET, and TEAMS_TENANT_ID env vars

Configuration in config.yaml:
    platforms:
      teams:
        enabled: true
        extra:
          client_id: "your-client-id"      # or TEAMS_CLIENT_ID env var
          client_secret: "your-secret"      # or TEAMS_CLIENT_SECRET env var
          tenant_id: "your-tenant-id"       # or TEAMS_TENANT_ID env var
          port: 3978                        # or TEAMS_PORT env var
"""

import asyncio
import json
import logging
import os
from typing import Any, Dict, Optional

try:
    from aiohttp import web

    AIOHTTP_AVAILABLE = True
except ImportError:
    AIOHTTP_AVAILABLE = False
    web = None  # type: ignore[assignment]

try:
    from microsoft_teams.apps import App, ActivityContext
    from microsoft_teams.api import MessageActivity
    from microsoft_teams.api.activities.typing import TypingActivityInput
    from microsoft_teams.apps.http.adapter import (
        HttpMethod,
        HttpRequest,
        HttpResponse,
        HttpRouteHandler,
    )

    TEAMS_SDK_AVAILABLE = True
except ImportError:
    TEAMS_SDK_AVAILABLE = False
    App = None  # type: ignore[assignment,misc]
    ActivityContext = None  # type: ignore[assignment,misc]
    MessageActivity = None  # type: ignore[assignment,misc]
    TypingActivityInput = None  # type: ignore[assignment,misc]
    HttpMethod = str  # type: ignore[assignment,misc]
    HttpRequest = None  # type: ignore[assignment,misc]
    HttpResponse = None  # type: ignore[assignment,misc]
    HttpRouteHandler = None  # type: ignore[assignment,misc]

from gateway.config import Platform, PlatformConfig
from gateway.platforms.helpers import MessageDeduplicator
from gateway.platforms.base import (
    BasePlatformAdapter,
    MessageEvent,
    MessageType,
    SendResult,
    cache_image_from_url,
)

logger = logging.getLogger(__name__)

_DEFAULT_PORT = 3978
_WEBHOOK_PATH = "/api/messages"


class _AiohttpBridgeAdapter:
    """HttpServerAdapter that bridges the Teams SDK into an aiohttp server.

    Without a custom adapter, ``App()`` unconditionally imports fastapi/uvicorn
    and allocates a ``FastAPI()`` instance.  This bridge captures the SDK's
    route registrations and wires them into our own aiohttp ``Application``.
    """

    def __init__(self, aiohttp_app: "web.Application"):
        self._aiohttp_app = aiohttp_app

    def register_route(self, method: "HttpMethod", path: str, handler: "HttpRouteHandler") -> None:
        """Register an SDK route handler as an aiohttp route."""

        async def _aiohttp_handler(request: "web.Request") -> "web.Response":
            body = await request.json()
            headers = dict(request.headers)
            result: "HttpResponse" = await handler(HttpRequest(body=body, headers=headers))
            status = result.get("status", 200)
            resp_body = result.get("body")
            if resp_body is not None:
                return web.Response(
                    status=status,
                    body=json.dumps(resp_body),
                    content_type="application/json",
                )
            return web.Response(status=status)

        self._aiohttp_app.router.add_route(method, path, _aiohttp_handler)

    def serve_static(self, path: str, directory: str) -> None:
        pass

    async def start(self, port: int) -> None:
        raise NotImplementedError("aiohttp server is managed by the adapter")

    async def stop(self) -> None:
        pass


def check_teams_requirements() -> bool:
    """Check if Teams dependencies are available and configured."""
    if not TEAMS_SDK_AVAILABLE:
        return False
    if not AIOHTTP_AVAILABLE:
        return False
    client_id = os.getenv("TEAMS_CLIENT_ID", "")
    client_secret = os.getenv("TEAMS_CLIENT_SECRET", "")
    tenant_id = os.getenv("TEAMS_TENANT_ID", "")
    if not client_id or not client_secret or not tenant_id:
        return False
    return True


class TeamsAdapter(BasePlatformAdapter):
    """Microsoft Teams adapter using the microsoft-teams-apps SDK."""

    MAX_MESSAGE_LENGTH = 28000  # Teams text message limit (~28 KB)

    def __init__(self, config: PlatformConfig):
        super().__init__(config, Platform.TEAMS)
        extra = config.extra or {}
        self._client_id = extra.get("client_id") or os.getenv("TEAMS_CLIENT_ID", "")
        self._client_secret = extra.get("client_secret") or os.getenv("TEAMS_CLIENT_SECRET", "")
        self._tenant_id = extra.get("tenant_id") or os.getenv("TEAMS_TENANT_ID", "")
        self._port = int(extra.get("port") or os.getenv("TEAMS_PORT", str(_DEFAULT_PORT)))
        self._app: Optional["App"] = None
        self._runner: Optional["web.AppRunner"] = None
        self._dedup = MessageDeduplicator(max_size=1000)

    async def connect(self) -> bool:
        if not TEAMS_SDK_AVAILABLE:
            self._set_fatal_error(
                "MISSING_SDK",
                "microsoft-teams-apps not installed. Run: pip install microsoft-teams-apps",
                retryable=False,
            )
            return False

        if not AIOHTTP_AVAILABLE:
            self._set_fatal_error(
                "MISSING_SDK",
                "aiohttp not installed. Run: pip install aiohttp",
                retryable=False,
            )
            return False

        if not self._client_id or not self._client_secret or not self._tenant_id:
            self._set_fatal_error(
                "MISSING_CREDENTIALS",
                "TEAMS_CLIENT_ID, TEAMS_CLIENT_SECRET, and TEAMS_TENANT_ID are all required",
                retryable=False,
            )
            return False

        try:
            # Set up aiohttp app first — the bridge adapter wires SDK routes into it
            aiohttp_app = web.Application()
            aiohttp_app.router.add_get("/health", lambda _: web.Response(text="ok"))

            self._app = App(
                client_id=self._client_id,
                client_secret=self._client_secret,
                tenant_id=self._tenant_id,
                http_server_adapter=_AiohttpBridgeAdapter(aiohttp_app),
            )

            # Register message handler before initialize()
            @self._app.on_message
            async def _handle_message(ctx: ActivityContext[MessageActivity]):
                await self._on_message(ctx)

            # initialize() calls register_route() on the bridge, which adds
            # POST /api/messages to aiohttp_app automatically
            await self._app.initialize()

            self._runner = web.AppRunner(aiohttp_app)
            await self._runner.setup()
            site = web.TCPSite(self._runner, "0.0.0.0", self._port)
            await site.start()

            self._running = True
            self._mark_connected()
            logger.info(
                "[teams] Webhook server listening on 0.0.0.0:%d%s",
                self._port,
                _WEBHOOK_PATH,
            )
            return True

        except Exception as e:
            self._set_fatal_error(
                "CONNECT_FAILED",
                f"Teams connection failed: {e}",
                retryable=True,
            )
            logger.error("[teams] Failed to connect: %s", e)
            return False

    async def disconnect(self) -> None:
        self._running = False
        if self._runner:
            await self._runner.cleanup()
            self._runner = None
        self._app = None
        self._mark_disconnected()
        logger.info("[teams] Disconnected")

    async def _on_message(self, ctx: ActivityContext[MessageActivity]) -> None:
        """Process an incoming Teams message and dispatch to the gateway."""
        activity = ctx.activity

        # Self-message filter
        bot_id = self._app.id if self._app else None
        if bot_id and getattr(activity.from_, "id", None) == bot_id:
            return

        # Deduplication
        msg_id = getattr(activity, "id", None)
        if msg_id and self._dedup.is_duplicate(msg_id):
            return

        # Extract text — strip bot @mentions
        text = ""
        if hasattr(activity, "text") and activity.text:
            text = activity.text
        # Strip <at>BotName</at> HTML tags that Teams prepends for @mentions
        if "<at>" in text:
            import re
            text = re.sub(r"<at>[^<]*</at>\s*", "", text).strip()

        # Determine chat type from conversation
        conv = activity.conversation
        conv_type = getattr(conv, "conversation_type", None) or ""
        if conv_type == "personal":
            chat_type = "dm"
        elif conv_type == "groupChat":
            chat_type = "group"
        elif conv_type == "channel":
            chat_type = "channel"
        else:
            chat_type = "dm"

        # Build source
        from_account = activity.from_
        user_id = getattr(from_account, "aad_object_id", None) or getattr(from_account, "id", "")
        user_name = getattr(from_account, "name", None) or ""

        source = self.build_source(
            chat_id=conv.id,
            chat_name=getattr(conv, "name", None) or "",
            chat_type=chat_type,
            user_id=str(user_id),
            user_name=user_name,
            guild_id=getattr(conv, "tenant_id", None) or self._tenant_id,
        )

        # Handle image attachments
        media_urls = []
        media_types = []
        for att in getattr(activity, "attachments", None) or []:
            content_url = getattr(att, "content_url", None)
            content_type = getattr(att, "content_type", None) or ""
            if content_url and content_type.startswith("image/"):
                try:
                    cached = await cache_image_from_url(content_url)
                    if cached:
                        media_urls.append(cached)
                        media_types.append(content_type)
                except Exception as e:
                    logger.warning("[teams] Failed to cache image attachment: %s", e)

        msg_type = MessageType.PHOTO if media_urls else MessageType.TEXT

        event = MessageEvent(
            text=text,
            source=source,
            message_type=msg_type,
            media_urls=media_urls,
            media_types=media_types,
            message_id=msg_id,
        )
        await self.handle_message(event)

    async def send(
        self,
        chat_id: str,
        content: str,
        reply_to: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> SendResult:
        if not self._app:
            return SendResult(success=False, error="Teams app not initialized")

        formatted = self.format_message(content)
        chunks = self.truncate_message(formatted)
        last_message_id = None

        for chunk in chunks:
            try:
                result = await self._app.send(chat_id, chunk)
                last_message_id = getattr(result, "id", None)
            except Exception as e:
                return SendResult(success=False, error=str(e), retryable=True)

        return SendResult(success=True, message_id=last_message_id)

    async def send_typing(self, chat_id: str, metadata: Optional[Dict[str, Any]] = None) -> None:
        if not self._app:
            return
        try:
            await self._app.send(chat_id, TypingActivityInput())
        except Exception:
            pass

    async def send_image(
        self,
        chat_id: str,
        image_url: str,
        caption: Optional[str] = None,
        reply_to: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> SendResult:
        # Teams: embed image as markdown
        text = f"![image]({image_url})"
        if caption:
            text = f"{caption}\n\n{text}"
        return await self.send(chat_id, text, reply_to=reply_to, metadata=metadata)

    async def get_chat_info(self, chat_id: str) -> dict:
        return {"name": chat_id, "type": "unknown", "chat_id": chat_id}
