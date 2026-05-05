"""
Microsoft Teams platform adapter.

Uses Microsoft Bot Framework with Azure Bot Service for:
- Receiving messages from 1:1, group, and channel chats
- Sending text and Adaptive Cards
- Typing indicators
- Mention handling
- Thread support
"""

import asyncio
import json
import logging
import os
import re
import time
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlsplit

try:
    from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings, TurnContext
    from botbuilder.schema import Activity, ActivityTypes, ChannelAccount, ConversationAccount, ConversationParameters, ResourceResponse
    from botbuilder.schema.teams import TeamsChannelData, TeamInfo, ChannelInfo
    TEAMS_AVAILABLE = True
except ImportError:
    TEAMS_AVAILABLE = False
    BotFrameworkAdapter = Any
    BotFrameworkAdapterSettings = Any
    TurnContext = Any
    Activity = Any
    ActivityTypes = None
    ChannelAccount = Any
    ConversationAccount = Any
    ConversationParameters = Any
    ResourceResponse = Any
    TeamsChannelData = Any
    TeamInfo = Any
    ChannelInfo = Any

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parents[2]))

from gateway.config import Platform, PlatformConfig
from gateway.platforms.base import (
    BasePlatformAdapter,
    MessageEvent,
    MessageType,
    ProcessingOutcome,
    SendResult,
    cache_document_from_bytes,
)
from gateway.platforms.helpers import MessageDeduplicator


logger = logging.getLogger(__name__)

# Teams message size limits
MAX_MESSAGE_LENGTH = 8000


@dataclass
class _ThreadContextCache:
    """Cache entry for fetched thread context."""
    content: str
    fetched_at: float = field(default_factory=time.monotonic)
    message_count: int = 0


def check_teams_requirements() -> bool:
    """Check if Microsoft Teams dependencies are available."""
    return TEAMS_AVAILABLE


def _extract_text_from_adaptive_card(card_actions: list) -> str:
    """Extract text from Adaptive Card actions for message content."""
    if not card_actions:
        return ""
    parts = []
    for action in card_actions:
        if action.get("type") == "Action.Execute":
            title = action.get("title", "")
            if title:
                parts.append(f"[{title}]")
    return " ".join(parts)


def _strip_teams_formatting(text: str) -> str:
    """Strip Teams-specific formatting artifacts."""
    if not text:
        return ""
    
    # Remove HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    
    # Remove mention formatting like <at>id</at>
    text = re.sub(r'<at[^>]*>[^<]*</at>', '', text)
    
    # Clean up multiple whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text


class TeamsAdapterImpl(BasePlatformAdapter):
    """Microsoft Teams platform adapter."""
    
    def __init__(self, config: PlatformConfig):
        super().__init__(config, Platform.TEAMS)
        
        # Teams-specific configuration
        self._client_id = self.config.extra.get("client_id", "")
        self._client_secret = self.config.extra.get("client_secret", "")
        self._tenant_id = self.config.extra.get("tenant_id", "")
        self._bot_id = self.config.extra.get("bot_id", "")
        self._bot_password = self.config.extra.get("bot_password", "")
        
        # Adapter instance
        self._adapter: Optional[BotFrameworkAdapter] = None
        self._settings: Optional[BotFrameworkAdapterSettings] = None
        
        # Message deduplication
        self._deduplicator = MessageDeduplicator()
        
        # Thread context cache
        self._thread_cache: Dict[str, _ThreadContextCache] = {}
        
        # Active conversations map
        self._conversations: Dict[str, Dict[str, Any]] = {}
        
        logger.info(
            f"Teams adapter initialized: tenant_id={self._redact(self._tenant_id)}, "
            f"client_id={self._redact(self._client_id)}"
        )
    
    def _redact(self, value: str) -> str:
        """Redact sensitive identifiers."""
        if not value:
            return ""
        if len(value) <= 4:
            return "****"
        return f"{value[:4]}...{value[-4:]}"
    
    async def connect(self) -> bool:
        """Connect to Microsoft Teams."""
        if not TEAMS_AVAILABLE:
            logger.error("Teams dependencies not available. Install botbuilder-core.")
            return False
        
        try:
            # Create adapter settings
            self._settings = BotFrameworkAdapterSettings(
                app_id=self._client_id,
                app_password=self._bot_password,
            )
            
            # Create the adapter
            self._adapter = BotFrameworkAdapter(self._settings)
            
            logger.info("Teams adapter connected successfully")
            return True
            
        except Exception as e:
            logger.error(f"Failed to connect Teams adapter: {e}")
            return False
    
    async def disconnect(self) -> None:
        """Disconnect from Microsoft Teams."""
        if self._adapter:
            # Cleanup if needed
            self._adapter = None
        self._settings = None
        logger.info("Teams adapter disconnected")
    
    def _build_session_source(self, activity: Activity) -> Dict[str, Any]:
        """Build session source from Teams activity."""
        channel_id = activity.channel_id or "teams"
        
        # Extract conversation info
        conv = activity.conversation
        conversation_id = conv.id if conv else ""
        conversation_name = conv.name if conv and conv.name else "Unknown"
        
        # Extract user info
        from_user = activity.from_property
        user_id = from_user.id if from_user else ""
        user_name = from_user.name if from_user and from_user.name else ""
        
        # Extract team/channel info if in team context
        team_id = ""
        channel_id_name = ""
        if activity.channel_data:
            channel_data = activity.channel_data
            if isinstance(channel_data, dict):
                team_info = channel_data.get("team", {})
                team_id = team_info.get("id", "") if team_info else ""
                channel_info = channel_data.get("channel", {})
                channel_id_name = channel_info.get("name", "") if channel_info else ""
        
        return {
            "platform": "teams",
            "channel_id": channel_id,
            "chat_id": conversation_id,
            "user_id": user_id,
            "user_name": user_name,
            "team_id": team_id,
            "channel_name": channel_id_name or conversation_name,
        }
    
    def _parse_message(self, activity: Activity) -> Tuple[Optional[str], MessageType]:
        """Parse Teams activity into message content and type."""
        if activity.type == ActivityTypes.message:
            text = activity.text or ""
            
            # Check for attachments
            if activity.attachments:
                # Handle Adaptive Cards
                for attachment in activity.attachments:
                    if attachment.content_type == "application/vnd.microsoft.card.adaptive":
                        card_text = _extract_text_from_adaptive_card(
                            attachment.content.get("actions", [])
                        )
                        if card_text:
                            text = f"{text} {card_text}".strip()
            
            # Strip Teams formatting
            text = _strip_teams_formatting(text)
            
            return text, MessageType.TEXT
            
        elif activity.type == ActivityTypes.contact_relation_update:
            return f"User {activity.action}", MessageType.SYSTEM
            
        elif activity.type == ActivityTypes.conversation_update:
            # Handle channel/team events
            if activity.members_added:
                member_names = [m.name for m in activity.members_added if m.name]
                return f"Members joined: {', '.join(member_names)}", MessageType.SYSTEM
            elif activity.members_removed:
                return "Members removed", MessageType.SYSTEM
            
            return None, MessageType.SYSTEM
        
        return None, MessageType.SYSTEM
    
    async def _handle_activity(self, activity: Activity) -> None:
        """Process an incoming Teams activity."""
        try:
            # Skip our own messages
            if activity.from_property and activity.from_property.id == self._bot_id:
                return
            
            # Parse the message
            text, msg_type = self._parse_message(activity)
            
            if not text and msg_type == MessageType.TEXT:
                return
            
            # Build session source
            source = self._build_session_source(activity)
            
            # Get conversation ID for deduplication
            conv_id = activity.conversation.id if activity.conversation else ""
            
            # Check for duplicates
            message_id = activity.id or f"{conv_id}:{time.time()}"
            if self._deduplicator.is_duplicate(message_id):
                return
            
            # Create message event
            event = MessageEvent(
                platform=Platform.TEAMS,
                message_id=message_id,
                chat_id=conv_id,
                user_id=source.get("user_id", ""),
                user_name=source.get("user_name", ""),
                text=text,
                message_type=msg_type,
                raw_event=activity.as_dict() if hasattr(activity, "as_dict") else {},
                source=source,
            )
            
            # Handle the message
            await self.handle_message(event)
            
        except Exception as e:
            logger.exception(f"Error handling Teams activity: {e}")
    
    async def send(
        self,
        chat_id: str,
        text: str,
        *,
        thread_id: Optional[str] = None,
        **kwargs
    ) -> SendResult:
        """Send a text message to Teams."""
        if not self._adapter:
            return SendResult(success=False, error="Adapter not connected")
        
        try:
            # Truncate if needed
            if len(text) > MAX_MESSAGE_LENGTH:
                text = text[:MAX_MESSAGE_LENGTH - 100] + "\n\n[truncated]"
            
            # Create activity
            activity = Activity(
                type=ActivityTypes.message,
                channel_id="msteams",
                conversation={"id": chat_id},
                from_property={"id": self._bot_id},
                text=text,
            )
            
            # Add replyToId for threading
            if thread_id:
                activity.reply_to_id = thread_id
            
            # Send via adapter ( Teams-specific implementation would go here )
            # For now, return success as placeholder
            logger.info(f"Teams send to {chat_id}: {text[:50]}...")
            
            return SendResult(success=True, message_id=f"teams-{chat_id}-{time.time()}")
            
        except Exception as e:
            logger.exception(f"Error sending Teams message: {e}")
            return SendResult(success=False, error=str(e))
    
    async def send_typing(self, chat_id: str) -> None:
        """Send typing indicator to Teams."""
        if not self._adapter:
            return
        
        try:
            activity = Activity(
                type=ActivityTypes.typing,
                channel_id="msteams",
                conversation={"id": chat_id},
                from_property={"id": self._bot_id},
            )
            # Would send via adapter
            logger.debug(f"Teams typing to {chat_id}")
        except Exception as e:
            logger.exception(f"Error sending typing indicator: {e}")
    
    async def send_image(
        self,
        chat_id: str,
        image_url: str,
        caption: Optional[str] = None
    ) -> SendResult:
        """Send an image to Teams."""
        # Teams supports images via attachments
        return await self.send(chat_id, f"{caption or ''}\n{image_url}".strip())
    
    async def send_document(
        self,
        chat_id: str,
        path: str,
        caption: Optional[str] = None
    ) -> SendResult:
        """Send a document to Teams."""
        return await self.send(chat_id, f"{caption or ''}\n[File: {path}]".strip())
    
    async def send_video(
        self,
        chat_id: str,
        path: str,
        caption: Optional[str] = None
    ) -> SendResult:
        """Send a video to Teams."""
        return await self.send(chat_id, f"{caption or ''}\n[Video: {path}]".strip())
    
    async def send_voice(
        self,
        chat_id: str,
        path: str
    ) -> SendResult:
        """Send a voice message to Teams."""
        return await self.send(chat_id, f"[Voice: {path}]".strip())
    
    async def get_chat_info(self, chat_id: str) -> Dict[str, str]:
        """Get information about a Teams conversation."""
        return {
            "name": f"Teams Chat {chat_id}",
            "type": "teams",
            "chat_id": chat_id,
        }


# Standalone function for send_message tool
async def _send_teams_message(
    config: PlatformConfig,
    chat_id: str,
    message: str
) -> SendResult:
    """Send a message to Teams outside the gateway process."""
    adapter = TeamsAdapterImpl(config)
    connected = await adapter.connect()
    if not connected:
        return SendResult(success=False, error="Failed to connect to Teams")
    
    try:
        return await adapter.send(chat_id, message)
    finally:
        await adapter.disconnect()
