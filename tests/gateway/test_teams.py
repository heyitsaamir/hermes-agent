"""Tests for the Microsoft Teams platform adapter."""

import asyncio
import os
import sys
import types
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from gateway.config import Platform, PlatformConfig, HomeChannel


# ---------------------------------------------------------------------------
# SDK Mock — install in sys.modules before importing the adapter
# ---------------------------------------------------------------------------

def _ensure_teams_mock():
    """Install a teams SDK mock in sys.modules if the real package isn't present."""
    if "microsoft_teams" in sys.modules and hasattr(sys.modules["microsoft_teams"], "__file__"):
        return

    # Build the module hierarchy
    microsoft_teams = types.ModuleType("microsoft_teams")
    microsoft_teams_apps = types.ModuleType("microsoft_teams.apps")
    microsoft_teams_api = types.ModuleType("microsoft_teams.api")
    microsoft_teams_api_activities = types.ModuleType("microsoft_teams.api.activities")
    microsoft_teams_api_activities_typing = types.ModuleType("microsoft_teams.api.activities.typing")
    microsoft_teams_apps_http = types.ModuleType("microsoft_teams.apps.http")
    microsoft_teams_apps_http_adapter = types.ModuleType("microsoft_teams.apps.http.adapter")

    # App class mock
    class MockApp:
        def __init__(self, **kwargs):
            self._client_id = kwargs.get("client_id")
            self.server = MagicMock()
            self.server.handle_request = AsyncMock(return_value={"status": 200, "body": None})
            self.credentials = MagicMock()
            self.credentials.client_id = self._client_id

        @property
        def id(self):
            return self._client_id

        def on_message(self, func):
            self._message_handler = func
            return func

        async def initialize(self):
            pass

        async def send(self, conversation_id, activity):
            result = MagicMock()
            result.id = "sent-activity-id"
            return result

        async def start(self, port=3978):
            pass

        async def stop(self):
            pass

    microsoft_teams_apps.App = MockApp
    microsoft_teams_apps.ActivityContext = MagicMock

    # MessageActivity mock
    microsoft_teams_api.MessageActivity = MagicMock

    # TypingActivityInput mock
    class MockTypingActivityInput:
        pass

    microsoft_teams_api_activities_typing.TypingActivityInput = MockTypingActivityInput

    # HttpRequest TypedDict mock
    def HttpRequest(body=None, headers=None):
        return {"body": body, "headers": headers}

    # HttpResponse TypedDict mock
    HttpResponse = dict

    # HttpMethod is just a string literal type
    HttpMethod = str

    # HttpRouteHandler is a callable protocol
    from typing import Callable
    HttpRouteHandler = Callable

    microsoft_teams_apps_http_adapter.HttpRequest = HttpRequest
    microsoft_teams_apps_http_adapter.HttpResponse = HttpResponse
    microsoft_teams_apps_http_adapter.HttpMethod = HttpMethod
    microsoft_teams_apps_http_adapter.HttpRouteHandler = HttpRouteHandler

    # Wire the hierarchy
    for name, mod in {
        "microsoft_teams": microsoft_teams,
        "microsoft_teams.apps": microsoft_teams_apps,
        "microsoft_teams.api": microsoft_teams_api,
        "microsoft_teams.api.activities": microsoft_teams_api_activities,
        "microsoft_teams.api.activities.typing": microsoft_teams_api_activities_typing,
        "microsoft_teams.apps.http": microsoft_teams_apps_http,
        "microsoft_teams.apps.http.adapter": microsoft_teams_apps_http_adapter,
    }.items():
        sys.modules.setdefault(name, mod)


_ensure_teams_mock()

# Now safe to import the adapter
import gateway.platforms.teams as _teams_mod

_teams_mod.TEAMS_SDK_AVAILABLE = True
_teams_mod.AIOHTTP_AVAILABLE = True

from gateway.platforms.teams import TeamsAdapter, check_teams_requirements


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_config(**extra):
    return PlatformConfig(enabled=True, extra=extra)


# ---------------------------------------------------------------------------
# Tests: Requirements
# ---------------------------------------------------------------------------

class TestTeamsRequirements:
    def test_returns_false_when_sdk_missing(self, monkeypatch):
        monkeypatch.setattr(_teams_mod, "TEAMS_SDK_AVAILABLE", False)
        assert check_teams_requirements() is False

    def test_returns_false_when_aiohttp_missing(self, monkeypatch):
        monkeypatch.setattr(_teams_mod, "AIOHTTP_AVAILABLE", False)
        assert check_teams_requirements() is False

    def test_returns_false_when_env_vars_missing(self, monkeypatch):
        monkeypatch.delenv("TEAMS_CLIENT_ID", raising=False)
        monkeypatch.delenv("TEAMS_CLIENT_SECRET", raising=False)
        monkeypatch.delenv("TEAMS_TENANT_ID", raising=False)
        assert check_teams_requirements() is False

    def test_returns_false_when_tenant_id_missing(self, monkeypatch):
        monkeypatch.setenv("TEAMS_CLIENT_ID", "test-id")
        monkeypatch.setenv("TEAMS_CLIENT_SECRET", "test-secret")
        monkeypatch.delenv("TEAMS_TENANT_ID", raising=False)
        assert check_teams_requirements() is False

    def test_returns_true_when_all_available(self, monkeypatch):
        monkeypatch.setattr(_teams_mod, "TEAMS_SDK_AVAILABLE", True)
        monkeypatch.setattr(_teams_mod, "AIOHTTP_AVAILABLE", True)
        monkeypatch.setenv("TEAMS_CLIENT_ID", "test-id")
        monkeypatch.setenv("TEAMS_CLIENT_SECRET", "test-secret")
        monkeypatch.setenv("TEAMS_TENANT_ID", "test-tenant")
        assert check_teams_requirements() is True


# ---------------------------------------------------------------------------
# Tests: Adapter Init
# ---------------------------------------------------------------------------

class TestTeamsAdapterInit:
    def test_reads_config_from_extra(self):
        config = _make_config(
            client_id="cfg-id",
            client_secret="cfg-secret",
            tenant_id="cfg-tenant",
        )
        adapter = TeamsAdapter(config)
        assert adapter._client_id == "cfg-id"
        assert adapter._client_secret == "cfg-secret"
        assert adapter._tenant_id == "cfg-tenant"

    def test_falls_back_to_env_vars(self, monkeypatch):
        monkeypatch.setenv("TEAMS_CLIENT_ID", "env-id")
        monkeypatch.setenv("TEAMS_CLIENT_SECRET", "env-secret")
        monkeypatch.setenv("TEAMS_TENANT_ID", "env-tenant")
        config = _make_config()
        adapter = TeamsAdapter(config)
        assert adapter._client_id == "env-id"
        assert adapter._client_secret == "env-secret"
        assert adapter._tenant_id == "env-tenant"

    def test_default_port(self):
        config = _make_config(client_id="id", client_secret="secret", tenant_id="tenant")
        adapter = TeamsAdapter(config)
        assert adapter._port == 3978

    def test_custom_port_from_extra(self):
        config = _make_config(client_id="id", client_secret="secret", tenant_id="tenant", port=4000)
        adapter = TeamsAdapter(config)
        assert adapter._port == 4000

    def test_custom_port_from_env(self, monkeypatch):
        monkeypatch.setenv("TEAMS_PORT", "5000")
        config = _make_config(client_id="id", client_secret="secret", tenant_id="tenant")
        adapter = TeamsAdapter(config)
        assert adapter._port == 5000


# ---------------------------------------------------------------------------
# Tests: Config Integration
# ---------------------------------------------------------------------------

class TestTeamsConfig:
    def test_platform_enum_exists(self):
        assert Platform.TEAMS.value == "teams"

    def test_env_overrides_loads_teams(self, monkeypatch):
        monkeypatch.setenv("TEAMS_CLIENT_ID", "test-id")
        monkeypatch.setenv("TEAMS_CLIENT_SECRET", "test-secret")
        monkeypatch.setenv("TEAMS_TENANT_ID", "test-tenant")

        from gateway.config import GatewayConfig
        config = GatewayConfig()
        # Simulate _apply_env_overrides by checking the env-based loading
        from gateway.config import _apply_env_overrides
        _apply_env_overrides(config)
        assert Platform.TEAMS in config.platforms
        assert config.platforms[Platform.TEAMS].enabled is True
        assert config.platforms[Platform.TEAMS].extra["client_id"] == "test-id"
        assert config.platforms[Platform.TEAMS].extra["tenant_id"] == "test-tenant"

    def test_env_overrides_skips_without_tenant_id(self, monkeypatch):
        monkeypatch.setenv("TEAMS_CLIENT_ID", "test-id")
        monkeypatch.setenv("TEAMS_CLIENT_SECRET", "test-secret")
        monkeypatch.delenv("TEAMS_TENANT_ID", raising=False)

        from gateway.config import GatewayConfig, _apply_env_overrides
        config = GatewayConfig()
        _apply_env_overrides(config)
        assert Platform.TEAMS not in config.platforms

    def test_env_overrides_loads_home_channel(self, monkeypatch):
        monkeypatch.setenv("TEAMS_CLIENT_ID", "test-id")
        monkeypatch.setenv("TEAMS_CLIENT_SECRET", "test-secret")
        monkeypatch.setenv("TEAMS_TENANT_ID", "test-tenant")
        monkeypatch.setenv("TEAMS_HOME_CHANNEL", "19:abc@thread.v2")
        monkeypatch.setenv("TEAMS_HOME_CHANNEL_NAME", "General")

        from gateway.config import GatewayConfig, _apply_env_overrides
        config = GatewayConfig()
        _apply_env_overrides(config)
        hc = config.platforms[Platform.TEAMS].home_channel
        assert hc is not None
        assert hc.chat_id == "19:abc@thread.v2"
        assert hc.name == "General"

    def test_get_connected_platforms_includes_teams(self, monkeypatch):
        monkeypatch.setenv("TEAMS_CLIENT_ID", "test-id")
        monkeypatch.setenv("TEAMS_CLIENT_SECRET", "test-secret")
        monkeypatch.setenv("TEAMS_TENANT_ID", "test-tenant")

        from gateway.config import GatewayConfig, _apply_env_overrides
        config = GatewayConfig()
        _apply_env_overrides(config)
        connected = config.get_connected_platforms()
        assert Platform.TEAMS in connected


# ---------------------------------------------------------------------------
# Tests: Authorization Maps
# ---------------------------------------------------------------------------

class TestTeamsAuthorization:
    """Verify Platform.TEAMS is wired into all three authorization maps."""

    def _get_runner_maps(self):
        """Read the authorization maps from gateway/run.py source."""
        import gateway.run as run_mod
        source = open(run_mod.__file__).read()
        return source

    def test_platform_in_env_map(self):
        source = self._get_runner_maps()
        assert 'Platform.TEAMS: "TEAMS_ALLOWED_USERS"' in source

    def test_platform_in_allow_all_map(self):
        source = self._get_runner_maps()
        assert 'Platform.TEAMS: "TEAMS_ALLOW_ALL_USERS"' in source

    def test_platform_in_unauthorized_dm_map(self):
        source = self._get_runner_maps()
        assert 'Platform.TEAMS:' in source
        assert '"TEAMS_ALLOWED_USERS"' in source


# ---------------------------------------------------------------------------
# Tests: Routing (send_message, cron, toolsets)
# ---------------------------------------------------------------------------

class TestTeamsRouting:
    def test_platform_in_send_message_map(self):
        """send_message_tool.py platform_map includes teams."""
        import tools.send_message_tool as smt_mod
        source = open(smt_mod.__file__).read()
        assert '"teams": Platform.TEAMS' in source

    def test_platform_in_cron_map(self):
        """cron/scheduler.py platform_map includes teams."""
        import cron.scheduler as cron_mod
        source = open(cron_mod.__file__).read()
        assert '"teams": Platform.TEAMS' in source

    def test_toolset_exists(self):
        from toolsets import TOOLSETS
        assert "hermes-teams" in TOOLSETS

    def test_in_gateway_toolset(self):
        from toolsets import TOOLSETS
        gateway = TOOLSETS["hermes-gateway"]
        assert "hermes-teams" in gateway["includes"]


# ---------------------------------------------------------------------------
# Tests: Connect / Disconnect
# ---------------------------------------------------------------------------

class TestTeamsConnect:
    @pytest.mark.asyncio
    async def test_connect_fails_without_sdk(self, monkeypatch):
        monkeypatch.setattr(_teams_mod, "TEAMS_SDK_AVAILABLE", False)
        adapter = TeamsAdapter(_make_config(
            client_id="id", client_secret="secret", tenant_id="tenant",
        ))
        result = await adapter.connect()
        assert result is False

    @pytest.mark.asyncio
    async def test_connect_fails_without_credentials(self):
        adapter = TeamsAdapter(_make_config())
        adapter._client_id = ""
        adapter._client_secret = ""
        adapter._tenant_id = ""
        result = await adapter.connect()
        assert result is False

    @pytest.mark.asyncio
    async def test_disconnect_cleans_up(self):
        adapter = TeamsAdapter(_make_config(
            client_id="id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._running = True
        mock_runner = AsyncMock()
        adapter._runner = mock_runner
        adapter._app = MagicMock()

        await adapter.disconnect()
        assert adapter._running is False
        assert adapter._app is None
        assert adapter._runner is None
        mock_runner.cleanup.assert_awaited_once()


# ---------------------------------------------------------------------------
# Tests: Send
# ---------------------------------------------------------------------------

class TestTeamsSend:
    @pytest.mark.asyncio
    async def test_send_returns_error_without_app(self):
        adapter = TeamsAdapter(_make_config(
            client_id="id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._app = None
        result = await adapter.send("conv-id", "Hello")
        assert result.success is False
        assert "not initialized" in result.error

    @pytest.mark.asyncio
    async def test_send_calls_app_send(self):
        adapter = TeamsAdapter(_make_config(
            client_id="id", client_secret="secret", tenant_id="tenant",
        ))
        mock_result = MagicMock()
        mock_result.id = "msg-123"
        mock_app = MagicMock()
        mock_app.send = AsyncMock(return_value=mock_result)
        adapter._app = mock_app

        result = await adapter.send("conv-id", "Hello")
        assert result.success is True
        assert result.message_id == "msg-123"
        mock_app.send.assert_awaited_once_with("conv-id", "Hello")

    @pytest.mark.asyncio
    async def test_send_handles_error(self):
        adapter = TeamsAdapter(_make_config(
            client_id="id", client_secret="secret", tenant_id="tenant",
        ))
        mock_app = MagicMock()
        mock_app.send = AsyncMock(side_effect=Exception("Network error"))
        adapter._app = mock_app

        result = await adapter.send("conv-id", "Hello")
        assert result.success is False
        assert "Network error" in result.error

    @pytest.mark.asyncio
    async def test_send_typing(self):
        adapter = TeamsAdapter(_make_config(
            client_id="id", client_secret="secret", tenant_id="tenant",
        ))
        mock_app = MagicMock()
        mock_app.send = AsyncMock()
        adapter._app = mock_app

        await adapter.send_typing("conv-id")
        mock_app.send.assert_awaited_once()
        # Second arg should be TypingActivityInput instance
        call_args = mock_app.send.call_args
        assert call_args[0][0] == "conv-id"


# ---------------------------------------------------------------------------
# Tests: Message Handling
# ---------------------------------------------------------------------------

class TestTeamsMessageHandling:
    def _make_activity(
        self,
        *,
        text="Hello",
        from_id="user-123",
        from_aad_id="aad-456",
        from_name="Test User",
        conversation_id="19:abc@thread.v2",
        conversation_type="personal",
        tenant_id="tenant-789",
        activity_id="activity-001",
        attachments=None,
    ):
        """Build a mock MessageActivity."""
        activity = MagicMock()
        activity.text = text
        activity.id = activity_id
        activity.from_ = MagicMock()
        activity.from_.id = from_id
        activity.from_.aad_object_id = from_aad_id
        activity.from_.name = from_name
        activity.conversation = MagicMock()
        activity.conversation.id = conversation_id
        activity.conversation.conversation_type = conversation_type
        activity.conversation.name = "Test Chat"
        activity.conversation.tenant_id = tenant_id
        activity.attachments = attachments or []
        return activity

    def _make_ctx(self, activity):
        """Build a mock ActivityContext wrapping the activity."""
        ctx = MagicMock()
        ctx.activity = activity
        return ctx

    @pytest.mark.asyncio
    async def test_personal_message_creates_dm_event(self):
        adapter = TeamsAdapter(_make_config(
            client_id="bot-id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._app = MagicMock()
        adapter._app.id = "bot-id"
        adapter.handle_message = AsyncMock()

        activity = self._make_activity(conversation_type="personal")
        await adapter._on_message(self._make_ctx(activity))

        adapter.handle_message.assert_awaited_once()
        event = adapter.handle_message.call_args[0][0]
        assert event.source.chat_type == "dm"

    @pytest.mark.asyncio
    async def test_group_message_creates_group_event(self):
        adapter = TeamsAdapter(_make_config(
            client_id="bot-id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._app = MagicMock()
        adapter._app.id = "bot-id"
        adapter.handle_message = AsyncMock()

        activity = self._make_activity(conversation_type="groupChat")
        await adapter._on_message(self._make_ctx(activity))

        event = adapter.handle_message.call_args[0][0]
        assert event.source.chat_type == "group"

    @pytest.mark.asyncio
    async def test_channel_message_creates_channel_event(self):
        adapter = TeamsAdapter(_make_config(
            client_id="bot-id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._app = MagicMock()
        adapter._app.id = "bot-id"
        adapter.handle_message = AsyncMock()

        activity = self._make_activity(conversation_type="channel")
        await adapter._on_message(self._make_ctx(activity))

        event = adapter.handle_message.call_args[0][0]
        assert event.source.chat_type == "channel"

    @pytest.mark.asyncio
    async def test_user_id_uses_aad_object_id(self):
        adapter = TeamsAdapter(_make_config(
            client_id="bot-id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._app = MagicMock()
        adapter._app.id = "bot-id"
        adapter.handle_message = AsyncMock()

        activity = self._make_activity(from_aad_id="aad-stable-id", from_id="teams-id")
        await adapter._on_message(self._make_ctx(activity))

        event = adapter.handle_message.call_args[0][0]
        assert event.source.user_id == "aad-stable-id"

    @pytest.mark.asyncio
    async def test_self_message_filtered(self):
        adapter = TeamsAdapter(_make_config(
            client_id="bot-id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._app = MagicMock()
        adapter._app.id = "bot-id"
        adapter.handle_message = AsyncMock()

        activity = self._make_activity(from_id="bot-id")
        await adapter._on_message(self._make_ctx(activity))

        adapter.handle_message.assert_not_awaited()

    @pytest.mark.asyncio
    async def test_bot_mention_stripped_from_text(self):
        adapter = TeamsAdapter(_make_config(
            client_id="bot-id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._app = MagicMock()
        adapter._app.id = "bot-id"
        adapter.handle_message = AsyncMock()

        activity = self._make_activity(
            text="<at>Hermes</at> what is the weather?",
            from_id="user-id",
        )
        await adapter._on_message(self._make_ctx(activity))

        event = adapter.handle_message.call_args[0][0]
        assert event.text == "what is the weather?"

    @pytest.mark.asyncio
    async def test_deduplication(self):
        adapter = TeamsAdapter(_make_config(
            client_id="bot-id", client_secret="secret", tenant_id="tenant",
        ))
        adapter._app = MagicMock()
        adapter._app.id = "bot-id"
        adapter.handle_message = AsyncMock()

        activity = self._make_activity(activity_id="msg-dup-001", from_id="user-id")
        ctx = self._make_ctx(activity)

        await adapter._on_message(ctx)
        await adapter._on_message(ctx)

        # Should only be called once
        assert adapter.handle_message.await_count == 1
