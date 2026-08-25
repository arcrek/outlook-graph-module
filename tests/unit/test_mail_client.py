"""Unit tests for Outlook Mail Client."""

from __future__ import annotations
import base64
import unittest
from unittest.mock import MagicMock, patch
from outlook_graph.auth.token_manager import OutlookTokenManager
from outlook_graph.mail.mail_client import OutlookMailClient
from outlook_graph.types.account import AccountCredentials
from outlook_graph.types.mail import GetMessagesOptions, WaitForEmailOptions


class TestMailClient(unittest.TestCase):
    """Test suite for OutlookMailClient operations."""

    def setUp(self) -> None:
        self.mock_account = AccountCredentials(
            email="user@hotmail.com",
            refresh_token="mock-rt",
            client_id="mock-cid",
        )
        self.mock_token_mgr = MagicMock(spec=OutlookTokenManager)
        self.mock_token_mgr.get_access_token.return_value = "mock-token"
        self.mail_client = OutlookMailClient(self.mock_token_mgr)

    def test_fetch_messages_and_map_to_outlook_message(self) -> None:
        mock_graph_response = {
            "value": [
                {
                    "id": "msg-1",
                    "subject": "Your verification code is 849201",
                    "bodyPreview": "849201 is your code",
                    "body": {
                        "contentType": "html",
                        "content": "<p>Your code is 849201</p>",
                    },
                    "from": {
                        "emailAddress": {
                            "name": "Discord",
                            "address": "noreply@discord.com",
                        }
                    },
                    "receivedDateTime": "2026-08-25T12:00:00Z",
                    "hasAttachments": False,
                    "isRead": False,
                }
            ]
        }

        with patch.object(
            self.mail_client.http_client,
            "get",
            return_value=mock_graph_response,
        ):
            messages = self.mail_client.get_messages(
                self.mock_account, GetMessagesOptions(top=1)
            )

            self.assertEqual(len(messages), 1)
            msg = messages[0]
            self.assertEqual(msg.id, "msg-1")
            self.assertEqual(msg.subject, "Your verification code is 849201")
            self.assertIsNotNone(msg.from_recipient)
            self.assertEqual(
                msg.from_recipient.email_address.address, "noreply@discord.com"
            )
            self.assertFalse(msg.is_read)

    def test_fetch_single_message_by_id(self) -> None:
        mock_msg = {
            "id": "msg-abc",
            "subject": "Security Alert",
            "body": {"contentType": "text", "content": "Alert info"},
            "receivedDateTime": "2026-08-25T12:00:00Z",
        }

        with patch.object(
            self.mail_client.http_client, "get", return_value=mock_msg
        ):
            msg = self.mail_client.get_message_by_id(
                self.mock_account, "msg-abc"
            )
            self.assertEqual(msg.id, "msg-abc")
            self.assertEqual(msg.subject, "Security Alert")

    def test_fetch_and_decode_base64_attachments(self) -> None:
        raw_text = "attachment secret content"
        b64_content = base64.b64encode(raw_text.encode("utf-8")).decode("ascii")

        mock_att_resp = {
            "value": [
                {
                    "id": "att-1",
                    "name": "invoice.pdf",
                    "contentType": "application/pdf",
                    "size": len(raw_text),
                    "isInline": False,
                    "contentBytes": b64_content,
                }
            ]
        }

        with patch.object(
            self.mail_client.http_client,
            "get",
            return_value=mock_att_resp,
        ):
            attachments = self.mail_client.get_message_attachments(
                self.mock_account, "msg-1"
            )
            self.assertEqual(len(attachments), 1)
            att = attachments[0]
            self.assertEqual(att.name, "invoice.pdf")
            self.assertEqual(att.content_type, "application/pdf")
            self.assertIsNotNone(att.data)
            self.assertEqual(att.data.decode("utf-8"), raw_text)

    def test_mark_as_read(self) -> None:
        with patch.object(
            self.mail_client.http_client, "patch"
        ) as mock_patch:
            self.mail_client.mark_as_read(self.mock_account, "msg-123", True)
            mock_patch.assert_called_once_with(
                self.mock_account,
                "https://graph.microsoft.com/v1.0/me/messages/msg-123",
                body={"isRead": True},
            )

    def test_delete_message(self) -> None:
        with patch.object(
            self.mail_client.http_client, "delete"
        ) as mock_delete:
            self.mail_client.delete_message(self.mock_account, "msg-123")
            mock_delete.assert_called_once_with(
                self.mock_account,
                "https://graph.microsoft.com/v1.0/me/messages/msg-123",
            )


if __name__ == "__main__":
    unittest.main()
