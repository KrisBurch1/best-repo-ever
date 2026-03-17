"""Salesforce Sales Cloud operations.

Provides helpers for common Sales Cloud objects:
  - Leads
  - Opportunities
  - Accounts
  - Contacts (shared with Service Cloud)
"""

from typing import Any, Dict, List, Optional

from .connection import SalesforceConnection


class SalesCloud:
    """High-level interface for Salesforce Sales Cloud.

    Args:
        connection: An authenticated :class:`SalesforceConnection` instance.
    """

    def __init__(self, connection: SalesforceConnection) -> None:
        self._conn = connection

    # ------------------------------------------------------------------
    # Leads
    # ------------------------------------------------------------------

    def create_lead(
        self,
        last_name: str,
        company: str,
        first_name: Optional[str] = None,
        email: Optional[str] = None,
        phone: Optional[str] = None,
        lead_source: Optional[str] = None,
        status: str = "Open - Not Contacted",
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """Create a new Lead record.

        Args:
            last_name:   Lead last name (required by Salesforce).
            company:     Lead company name (required by Salesforce).
            first_name:  Lead first name.
            email:       Lead email address.
            phone:       Lead phone number.
            lead_source: Source of the lead (Web, Phone, Email, etc.).
            status:      Lead status.
            **kwargs:    Any additional Lead fields.

        Returns:
            A dict with ``id``, ``success``, and ``errors`` keys.
        """
        payload: Dict[str, Any] = {
            "LastName": last_name,
            "Company": company,
            "Status": status,
        }
        if first_name:
            payload["FirstName"] = first_name
        if email:
            payload["Email"] = email
        if phone:
            payload["Phone"] = phone
        if lead_source:
            payload["LeadSource"] = lead_source
        payload.update(kwargs)
        return self._conn.client.Lead.create(payload)

    def get_lead(self, lead_id: str) -> Dict[str, Any]:
        """Retrieve a Lead record by its Salesforce ID.

        Args:
            lead_id: Salesforce Lead ID.

        Returns:
            A dict containing all Lead fields.
        """
        return self._conn.client.Lead.get(lead_id)

    def update_lead(self, lead_id: str, **fields: Any) -> int:
        """Update fields on an existing Lead.

        Args:
            lead_id:  Salesforce Lead ID.
            **fields: Field-name/value pairs to update.

        Returns:
            HTTP status code (204 on success).
        """
        return self._conn.client.Lead.update(lead_id, fields)

    def convert_lead(self, lead_id: str, **fields: Any) -> int:
        """Mark a Lead as converted.

        Args:
            lead_id:  Salesforce Lead ID.
            **fields: Additional fields to set during conversion (e.g.
                      ``ConvertedAccountId``, ``ConvertedContactId``).

        Returns:
            HTTP status code (204 on success).
        """
        return self.update_lead(lead_id, IsConverted=True, **fields)

    def search_leads(self, soql_where: str) -> List[Dict[str, Any]]:
        """Query Leads using a SOQL WHERE clause.

        Args:
            soql_where: A SOQL WHERE clause, e.g. ``"Status = 'New'"``.

        Returns:
            A list of Lead record dicts.
        """
        query = (
            f"SELECT Id, FirstName, LastName, Company, Email, Status "
            f"FROM Lead WHERE {soql_where}"
        )
        result = self._conn.client.query_all(query)
        return result.get("records", [])

    # ------------------------------------------------------------------
    # Opportunities
    # ------------------------------------------------------------------

    def create_opportunity(
        self,
        name: str,
        stage: str,
        close_date: str,
        account_id: Optional[str] = None,
        amount: Optional[float] = None,
        lead_source: Optional[str] = None,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """Create a new Opportunity record.

        Args:
            name:        Opportunity name (required by Salesforce).
            stage:       Sales stage (e.g. "Prospecting", "Closed Won").
            close_date:  Expected close date in ``YYYY-MM-DD`` format (required).
            account_id:  ID of the related Account.
            amount:      Expected revenue amount.
            lead_source: Source of the opportunity.
            **kwargs:    Any additional Opportunity fields.

        Returns:
            A dict with ``id``, ``success``, and ``errors`` keys.
        """
        payload: Dict[str, Any] = {
            "Name": name,
            "StageName": stage,
            "CloseDate": close_date,
        }
        if account_id:
            payload["AccountId"] = account_id
        if amount is not None:
            payload["Amount"] = amount
        if lead_source:
            payload["LeadSource"] = lead_source
        payload.update(kwargs)
        return self._conn.client.Opportunity.create(payload)

    def get_opportunity(self, opportunity_id: str) -> Dict[str, Any]:
        """Retrieve an Opportunity record by its Salesforce ID.

        Args:
            opportunity_id: Salesforce Opportunity ID.

        Returns:
            A dict containing all Opportunity fields.
        """
        return self._conn.client.Opportunity.get(opportunity_id)

    def update_opportunity(self, opportunity_id: str, **fields: Any) -> int:
        """Update fields on an existing Opportunity.

        Args:
            opportunity_id: Salesforce Opportunity ID.
            **fields:       Field-name/value pairs to update.

        Returns:
            HTTP status code (204 on success).
        """
        return self._conn.client.Opportunity.update(opportunity_id, fields)

    def close_opportunity_won(self, opportunity_id: str) -> int:
        """Set an Opportunity stage to 'Closed Won'.

        Args:
            opportunity_id: Salesforce Opportunity ID.

        Returns:
            HTTP status code (204 on success).
        """
        return self.update_opportunity(opportunity_id, StageName="Closed Won")

    def search_opportunities(self, soql_where: str) -> List[Dict[str, Any]]:
        """Query Opportunities using a SOQL WHERE clause.

        Args:
            soql_where: A SOQL WHERE clause, e.g.
                        ``"StageName = 'Prospecting'"``.

        Returns:
            A list of Opportunity record dicts.
        """
        query = (
            f"SELECT Id, Name, StageName, Amount, CloseDate "
            f"FROM Opportunity WHERE {soql_where}"
        )
        result = self._conn.client.query_all(query)
        return result.get("records", [])

    # ------------------------------------------------------------------
    # Accounts
    # ------------------------------------------------------------------

    def create_account(
        self,
        name: str,
        industry: Optional[str] = None,
        phone: Optional[str] = None,
        website: Optional[str] = None,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """Create a new Account record.

        Args:
            name:     Account name (required by Salesforce).
            industry: Account industry classification.
            phone:    Primary phone number.
            website:  Account website URL.
            **kwargs: Any additional Account fields.

        Returns:
            A dict with ``id``, ``success``, and ``errors`` keys.
        """
        payload: Dict[str, Any] = {"Name": name}
        if industry:
            payload["Industry"] = industry
        if phone:
            payload["Phone"] = phone
        if website:
            payload["Website"] = website
        payload.update(kwargs)
        return self._conn.client.Account.create(payload)

    def get_account(self, account_id: str) -> Dict[str, Any]:
        """Retrieve an Account record by its Salesforce ID.

        Args:
            account_id: Salesforce Account ID.

        Returns:
            A dict containing all Account fields.
        """
        return self._conn.client.Account.get(account_id)

    def update_account(self, account_id: str, **fields: Any) -> int:
        """Update fields on an existing Account.

        Args:
            account_id: Salesforce Account ID.
            **fields:   Field-name/value pairs to update.

        Returns:
            HTTP status code (204 on success).
        """
        return self._conn.client.Account.update(account_id, fields)

    def search_accounts(self, soql_where: str) -> List[Dict[str, Any]]:
        """Query Accounts using a SOQL WHERE clause.

        Args:
            soql_where: A SOQL WHERE clause, e.g. ``"Industry = 'Technology'"``.

        Returns:
            A list of Account record dicts.
        """
        query = (
            f"SELECT Id, Name, Industry, Phone, Website "
            f"FROM Account WHERE {soql_where}"
        )
        result = self._conn.client.query_all(query)
        return result.get("records", [])
