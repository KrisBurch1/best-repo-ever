"""Tests for salesforce/connection.py"""

import os
from unittest.mock import MagicMock, patch

import pytest

from salesforce.connection import SalesforceConnection


class TestSalesforceConnectionInit:
    def test_credentials_from_kwargs(self):
        conn = SalesforceConnection(
            username="user@example.com",
            password="secret",
            security_token="token123",
            domain="test",
        )
        assert conn._username == "user@example.com"
        assert conn._password == "secret"
        assert conn._security_token == "token123"
        assert conn._domain == "test"

    def test_credentials_from_env(self, monkeypatch):
        monkeypatch.setenv("SF_USERNAME", "env_user@example.com")
        monkeypatch.setenv("SF_PASSWORD", "env_pass")
        monkeypatch.setenv("SF_SECURITY_TOKEN", "env_token")
        monkeypatch.setenv("SF_DOMAIN", "test")

        conn = SalesforceConnection()
        assert conn._username == "env_user@example.com"
        assert conn._password == "env_pass"
        assert conn._security_token == "env_token"
        assert conn._domain == "test"

    def test_default_domain_is_login(self, monkeypatch):
        monkeypatch.delenv("SF_DOMAIN", raising=False)
        conn = SalesforceConnection()
        assert conn._domain == "login"

    def test_kwargs_override_env(self, monkeypatch):
        monkeypatch.setenv("SF_USERNAME", "env_user@example.com")
        conn = SalesforceConnection(username="kwarg_user@example.com")
        assert conn._username == "kwarg_user@example.com"

    def test_not_connected_initially(self):
        conn = SalesforceConnection(username="u", password="p")
        assert conn.is_connected() is False


class TestSalesforceConnectionAuthentication:
    @patch("salesforce.connection.Salesforce")
    def test_password_auth(self, mock_sf_class):
        mock_sf_instance = MagicMock()
        mock_sf_class.return_value = mock_sf_instance

        conn = SalesforceConnection(
            username="user@example.com",
            password="secret",
            security_token="token",
            domain="login",
        )
        client = conn.client

        mock_sf_class.assert_called_once_with(
            username="user@example.com",
            password="secret",
            security_token="token",
            domain="login",
        )
        assert client is mock_sf_instance

    @patch("salesforce.connection.SalesforceLogin")
    @patch("salesforce.connection.Salesforce")
    def test_connected_app_auth(self, mock_sf_class, mock_login):
        mock_login.return_value = ("session_id_123", "na1.salesforce.com")
        mock_sf_instance = MagicMock()
        mock_sf_class.return_value = mock_sf_instance

        conn = SalesforceConnection(
            username="user@example.com",
            password="secret",
            security_token="token",
            consumer_key="key123",
            consumer_secret="secret123",
        )
        client = conn.client

        mock_login.assert_called_once()
        mock_sf_class.assert_called_once_with(
            instance="na1.salesforce.com", session_id="session_id_123"
        )
        assert client is mock_sf_instance

    @patch("salesforce.connection.Salesforce")
    def test_client_cached_after_first_call(self, mock_sf_class):
        mock_sf_class.return_value = MagicMock()
        conn = SalesforceConnection(username="user@example.com", password="secret")

        _ = conn.client
        _ = conn.client

        mock_sf_class.assert_called_once()

    def test_password_auth_raises_on_missing_credentials(self):
        conn = SalesforceConnection()
        conn._username = None
        with pytest.raises(ValueError, match="SF_USERNAME"):
            conn.client

    def test_connected_app_auth_raises_on_missing_username(self):
        conn = SalesforceConnection(
            consumer_key="key", consumer_secret="secret"
        )
        conn._username = None
        with pytest.raises(ValueError, match="SF_USERNAME"):
            conn.client

    @patch("salesforce.connection.Salesforce")
    def test_is_connected_true_after_auth(self, mock_sf_class):
        mock_sf_class.return_value = MagicMock()
        conn = SalesforceConnection(username="user@example.com", password="secret")
        _ = conn.client
        assert conn.is_connected() is True

    @patch("salesforce.connection.Salesforce")
    def test_disconnect_clears_client(self, mock_sf_class):
        mock_sf_class.return_value = MagicMock()
        conn = SalesforceConnection(username="user@example.com", password="secret")
        _ = conn.client
        assert conn.is_connected() is True

        conn.disconnect()
        assert conn.is_connected() is False

    @patch("salesforce.connection.Salesforce")
    def test_reconnect_after_disconnect(self, mock_sf_class):
        mock_sf_class.return_value = MagicMock()
        conn = SalesforceConnection(username="user@example.com", password="secret")
        _ = conn.client
        conn.disconnect()
        _ = conn.client

        assert mock_sf_class.call_count == 2
