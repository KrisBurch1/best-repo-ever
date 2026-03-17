"""Core Salesforce authentication and connection management."""

import os
from typing import Optional

from simple_salesforce import Salesforce, SalesforceLogin
from simple_salesforce.exceptions import SalesforceAuthenticationFailed


class SalesforceConnection:
    """Manages authentication and connection to Salesforce.

    Supports username/password authentication as well as
    connected app OAuth2 (client credentials) authentication.

    Environment variables used when constructor parameters are not provided:
        SF_USERNAME       - Salesforce username
        SF_PASSWORD       - Salesforce password
        SF_SECURITY_TOKEN - Salesforce security token
        SF_CONSUMER_KEY   - Connected app consumer key (client ID)
        SF_CONSUMER_SECRET - Connected app consumer secret (client secret)
        SF_DOMAIN         - Login domain, e.g. "login" or "test" (default: "login")
    """

    def __init__(
        self,
        username: Optional[str] = None,
        password: Optional[str] = None,
        security_token: Optional[str] = None,
        consumer_key: Optional[str] = None,
        consumer_secret: Optional[str] = None,
        domain: Optional[str] = None,
    ) -> None:
        self._username = username or os.environ.get("SF_USERNAME")
        self._password = password or os.environ.get("SF_PASSWORD")
        self._security_token = security_token or os.environ.get("SF_SECURITY_TOKEN", "")
        self._consumer_key = consumer_key or os.environ.get("SF_CONSUMER_KEY")
        self._consumer_secret = consumer_secret or os.environ.get("SF_CONSUMER_SECRET")
        self._domain = domain or os.environ.get("SF_DOMAIN", "login")
        self._sf: Optional[Salesforce] = None

    @property
    def client(self) -> Salesforce:
        """Return the authenticated Salesforce client, connecting if necessary."""
        if self._sf is None:
            self._sf = self._authenticate()
        return self._sf

    def _authenticate(self) -> Salesforce:
        """Authenticate against Salesforce and return a client instance.

        Raises:
            SalesforceAuthenticationFailed: If authentication fails.
            ValueError: If required credentials are missing.
        """
        if self._consumer_key and self._consumer_secret:
            return self._authenticate_with_connected_app()
        return self._authenticate_with_password()

    def _authenticate_with_password(self) -> Salesforce:
        """Authenticate using username + password + security token."""
        if not self._username or not self._password:
            raise ValueError(
                "SF_USERNAME and SF_PASSWORD must be set for password authentication."
            )
        return Salesforce(
            username=self._username,
            password=self._password,
            security_token=self._security_token,
            domain=self._domain,
        )

    def _authenticate_with_connected_app(self) -> Salesforce:
        """Authenticate using a connected app (consumer key / secret)."""
        if not self._username or not self._password:
            raise ValueError(
                "SF_USERNAME and SF_PASSWORD must be set for connected app authentication."
            )
        session_id, instance = SalesforceLogin(
            username=self._username,
            password=self._password,
            security_token=self._security_token,
            consumer_key=self._consumer_key,
            consumer_secret=self._consumer_secret,
            domain=self._domain,
        )
        return Salesforce(instance=instance, session_id=session_id)

    def is_connected(self) -> bool:
        """Return True if a connection has been established."""
        return self._sf is not None

    def disconnect(self) -> None:
        """Drop the current connection so the next call will re-authenticate."""
        self._sf = None
