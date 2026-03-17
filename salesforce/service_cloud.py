"""Salesforce Service Cloud operations.

Provides helpers for common Service Cloud objects:
  - Cases
  - Contacts (shared with Sales Cloud)
  - Case Comments
"""

from typing import Any, Dict, List, Optional

from .connection import SalesforceConnection


class ServiceCloud:
    """High-level interface for Salesforce Service Cloud.

    Args:
        connection: An authenticated :class:`SalesforceConnection` instance.
    """

    def __init__(self, connection: SalesforceConnection) -> None:
        self._conn = connection

    # ------------------------------------------------------------------
    # Cases
    # ------------------------------------------------------------------

    def create_case(
        self,
        subject: str,
        description: Optional[str] = None,
        contact_id: Optional[str] = None,
        account_id: Optional[str] = None,
        priority: str = "Medium",
        status: str = "New",
        origin: str = "Web",
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """Create a new Case record.

        Args:
            subject:     Case subject line.
            description: Detailed description of the issue.
            contact_id:  ID of the related Contact.
            account_id:  ID of the related Account.
            priority:    Case priority (Low, Medium, High).
            status:      Case status (New, Working, Escalated, Closed).
            origin:      Case origin (Web, Phone, Email).
            **kwargs:    Any additional Case fields.

        Returns:
            A dict with ``id``, ``success``, and ``errors`` keys.
        """
        payload: Dict[str, Any] = {
            "Subject": subject,
            "Priority": priority,
            "Status": status,
            "Origin": origin,
        }
        if description:
            payload["Description"] = description
        if contact_id:
            payload["ContactId"] = contact_id
        if account_id:
            payload["AccountId"] = account_id
        payload.update(kwargs)
        return self._conn.client.Case.create(payload)

    def get_case(self, case_id: str) -> Dict[str, Any]:
        """Retrieve a Case record by its Salesforce ID.

        Args:
            case_id: Salesforce Case ID (15- or 18-character).

        Returns:
            A dict containing all Case fields.
        """
        return self._conn.client.Case.get(case_id)

    def update_case(self, case_id: str, **fields: Any) -> int:
        """Update fields on an existing Case.

        Args:
            case_id:  Salesforce Case ID.
            **fields: Field-name/value pairs to update.

        Returns:
            HTTP status code (204 on success).
        """
        return self._conn.client.Case.update(case_id, fields)

    def close_case(self, case_id: str) -> int:
        """Set a Case status to 'Closed'.

        Args:
            case_id: Salesforce Case ID.

        Returns:
            HTTP status code (204 on success).
        """
        return self.update_case(case_id, Status="Closed")

    def search_cases(self, soql_where: str) -> List[Dict[str, Any]]:
        """Query Cases using a SOQL WHERE clause.

        Args:
            soql_where: A SOQL WHERE clause, e.g.
                        ``"Status = 'New' AND Priority = 'High'"``.

        Returns:
            A list of Case record dicts.
        """
        query = f"SELECT Id, CaseNumber, Subject, Status, Priority FROM Case WHERE {soql_where}"
        result = self._conn.client.query_all(query)
        return result.get("records", [])

    # ------------------------------------------------------------------
    # Case Comments
    # ------------------------------------------------------------------

    def add_case_comment(
        self, case_id: str, body: str, is_published: bool = True
    ) -> Dict[str, Any]:
        """Add a CaseComment to an existing Case.

        Args:
            case_id:      Salesforce Case ID.
            body:         Comment text.
            is_published: Whether the comment is visible to the customer portal.

        Returns:
            A dict with ``id``, ``success``, and ``errors`` keys.
        """
        return self._conn.client.CaseComment.create(
            {"ParentId": case_id, "CommentBody": body, "IsPublished": is_published}
        )

    # ------------------------------------------------------------------
    # Contacts
    # ------------------------------------------------------------------

    def create_contact(
        self,
        last_name: str,
        first_name: Optional[str] = None,
        email: Optional[str] = None,
        phone: Optional[str] = None,
        account_id: Optional[str] = None,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """Create a Contact record.

        Args:
            last_name:  Contact last name (required by Salesforce).
            first_name: Contact first name.
            email:      Primary email address.
            phone:      Primary phone number.
            account_id: ID of the related Account.
            **kwargs:   Any additional Contact fields.

        Returns:
            A dict with ``id``, ``success``, and ``errors`` keys.
        """
        payload: Dict[str, Any] = {"LastName": last_name}
        if first_name:
            payload["FirstName"] = first_name
        if email:
            payload["Email"] = email
        if phone:
            payload["Phone"] = phone
        if account_id:
            payload["AccountId"] = account_id
        payload.update(kwargs)
        return self._conn.client.Contact.create(payload)

    def get_contact(self, contact_id: str) -> Dict[str, Any]:
        """Retrieve a Contact record by its Salesforce ID.

        Args:
            contact_id: Salesforce Contact ID.

        Returns:
            A dict containing all Contact fields.
        """
        return self._conn.client.Contact.get(contact_id)
