"""Tests for salesforce/sales_cloud.py"""

from unittest.mock import MagicMock

import pytest

from salesforce.connection import SalesforceConnection
from salesforce.sales_cloud import SalesCloud


@pytest.fixture()
def mock_conn():
    conn = MagicMock(spec=SalesforceConnection)
    conn.client = MagicMock()
    return conn


@pytest.fixture()
def sales(mock_conn):
    return SalesCloud(mock_conn)


class TestCreateLead:
    def test_required_fields(self, sales, mock_conn):
        mock_conn.client.Lead.create.return_value = {
            "id": "00Q1x",
            "success": True,
            "errors": [],
        }
        result = sales.create_lead("Doe", "Acme Corp")
        mock_conn.client.Lead.create.assert_called_once_with(
            {
                "LastName": "Doe",
                "Company": "Acme Corp",
                "Status": "Open - Not Contacted",
            }
        )
        assert result["success"] is True

    def test_optional_fields(self, sales, mock_conn):
        mock_conn.client.Lead.create.return_value = {"id": "00Q1x", "success": True, "errors": []}
        sales.create_lead(
            "Doe",
            "Acme Corp",
            first_name="Jane",
            email="jane@acme.com",
            phone="555-9999",
            lead_source="Web",
            status="Working - Contacted",
        )
        payload = mock_conn.client.Lead.create.call_args[0][0]
        assert payload["FirstName"] == "Jane"
        assert payload["Email"] == "jane@acme.com"
        assert payload["Phone"] == "555-9999"
        assert payload["LeadSource"] == "Web"
        assert payload["Status"] == "Working - Contacted"

    def test_extra_kwargs(self, sales, mock_conn):
        mock_conn.client.Lead.create.return_value = {"id": "00Q1x", "success": True, "errors": []}
        sales.create_lead("Doe", "Acme", Rating="Hot")
        payload = mock_conn.client.Lead.create.call_args[0][0]
        assert payload["Rating"] == "Hot"


class TestGetLead:
    def test_returns_lead(self, sales, mock_conn):
        mock_conn.client.Lead.get.return_value = {"Id": "00Q1x", "LastName": "Doe"}
        result = sales.get_lead("00Q1x")
        mock_conn.client.Lead.get.assert_called_once_with("00Q1x")
        assert result["LastName"] == "Doe"


class TestUpdateLead:
    def test_update_fields(self, sales, mock_conn):
        mock_conn.client.Lead.update.return_value = 204
        status_code = sales.update_lead("00Q1x", Status="Working - Contacted")
        mock_conn.client.Lead.update.assert_called_once_with(
            "00Q1x", {"Status": "Working - Contacted"}
        )
        assert status_code == 204


class TestConvertLead:
    def test_sets_is_converted(self, sales, mock_conn):
        mock_conn.client.Lead.update.return_value = 204
        sales.convert_lead("00Q1x")
        payload = mock_conn.client.Lead.update.call_args[0][1]
        assert payload["IsConverted"] is True

    def test_extra_conversion_fields(self, sales, mock_conn):
        mock_conn.client.Lead.update.return_value = 204
        sales.convert_lead("00Q1x", ConvertedAccountId="0011x")
        payload = mock_conn.client.Lead.update.call_args[0][1]
        assert payload["ConvertedAccountId"] == "0011x"


class TestSearchLeads:
    def test_query_construction(self, sales, mock_conn):
        mock_conn.client.query_all.return_value = {"records": []}
        sales.search_leads("Status = 'New'")
        mock_conn.client.query_all.assert_called_once_with(
            "SELECT Id, FirstName, LastName, Company, Email, Status "
            "FROM Lead WHERE Status = 'New'"
        )

    def test_returns_records(self, sales, mock_conn):
        records = [{"Id": "00Q1x", "LastName": "Doe"}]
        mock_conn.client.query_all.return_value = {"records": records}
        result = sales.search_leads("Status = 'New'")
        assert result == records


class TestCreateOpportunity:
    def test_required_fields(self, sales, mock_conn):
        mock_conn.client.Opportunity.create.return_value = {
            "id": "0061x",
            "success": True,
            "errors": [],
        }
        result = sales.create_opportunity("Big Deal", "Prospecting", "2026-12-31")
        mock_conn.client.Opportunity.create.assert_called_once_with(
            {
                "Name": "Big Deal",
                "StageName": "Prospecting",
                "CloseDate": "2026-12-31",
            }
        )
        assert result["success"] is True

    def test_optional_fields(self, sales, mock_conn):
        mock_conn.client.Opportunity.create.return_value = {
            "id": "0061x",
            "success": True,
            "errors": [],
        }
        sales.create_opportunity(
            "Big Deal",
            "Prospecting",
            "2026-12-31",
            account_id="0011x",
            amount=50000.0,
            lead_source="Web",
        )
        payload = mock_conn.client.Opportunity.create.call_args[0][0]
        assert payload["AccountId"] == "0011x"
        assert payload["Amount"] == 50000.0
        assert payload["LeadSource"] == "Web"


class TestGetOpportunity:
    def test_returns_opportunity(self, sales, mock_conn):
        mock_conn.client.Opportunity.get.return_value = {"Id": "0061x", "Name": "Big Deal"}
        result = sales.get_opportunity("0061x")
        mock_conn.client.Opportunity.get.assert_called_once_with("0061x")
        assert result["Name"] == "Big Deal"


class TestUpdateOpportunity:
    def test_update_fields(self, sales, mock_conn):
        mock_conn.client.Opportunity.update.return_value = 204
        sales.update_opportunity("0061x", StageName="Qualification")
        mock_conn.client.Opportunity.update.assert_called_once_with(
            "0061x", {"StageName": "Qualification"}
        )


class TestCloseOpportunityWon:
    def test_sets_stage_closed_won(self, sales, mock_conn):
        mock_conn.client.Opportunity.update.return_value = 204
        sales.close_opportunity_won("0061x")
        payload = mock_conn.client.Opportunity.update.call_args[0][1]
        assert payload["StageName"] == "Closed Won"


class TestSearchOpportunities:
    def test_query_construction(self, sales, mock_conn):
        mock_conn.client.query_all.return_value = {"records": []}
        sales.search_opportunities("StageName = 'Prospecting'")
        mock_conn.client.query_all.assert_called_once_with(
            "SELECT Id, Name, StageName, Amount, CloseDate "
            "FROM Opportunity WHERE StageName = 'Prospecting'"
        )

    def test_returns_records(self, sales, mock_conn):
        records = [{"Id": "0061x", "Name": "Big Deal"}]
        mock_conn.client.query_all.return_value = {"records": records}
        result = sales.search_opportunities("StageName = 'Prospecting'")
        assert result == records


class TestCreateAccount:
    def test_required_fields(self, sales, mock_conn):
        mock_conn.client.Account.create.return_value = {
            "id": "0011x",
            "success": True,
            "errors": [],
        }
        result = sales.create_account("Acme Corp")
        mock_conn.client.Account.create.assert_called_once_with({"Name": "Acme Corp"})
        assert result["success"] is True

    def test_optional_fields(self, sales, mock_conn):
        mock_conn.client.Account.create.return_value = {"id": "0011x", "success": True, "errors": []}
        sales.create_account(
            "Acme Corp",
            industry="Technology",
            phone="555-0000",
            website="https://acme.com",
        )
        payload = mock_conn.client.Account.create.call_args[0][0]
        assert payload["Industry"] == "Technology"
        assert payload["Phone"] == "555-0000"
        assert payload["Website"] == "https://acme.com"


class TestGetAccount:
    def test_returns_account(self, sales, mock_conn):
        mock_conn.client.Account.get.return_value = {"Id": "0011x", "Name": "Acme Corp"}
        result = sales.get_account("0011x")
        mock_conn.client.Account.get.assert_called_once_with("0011x")
        assert result["Name"] == "Acme Corp"


class TestUpdateAccount:
    def test_update_fields(self, sales, mock_conn):
        mock_conn.client.Account.update.return_value = 204
        sales.update_account("0011x", Industry="Finance")
        mock_conn.client.Account.update.assert_called_once_with(
            "0011x", {"Industry": "Finance"}
        )


class TestSearchAccounts:
    def test_query_construction(self, sales, mock_conn):
        mock_conn.client.query_all.return_value = {"records": []}
        sales.search_accounts("Industry = 'Technology'")
        mock_conn.client.query_all.assert_called_once_with(
            "SELECT Id, Name, Industry, Phone, Website "
            "FROM Account WHERE Industry = 'Technology'"
        )

    def test_returns_records(self, sales, mock_conn):
        records = [{"Id": "0011x", "Name": "Acme Corp"}]
        mock_conn.client.query_all.return_value = {"records": records}
        result = sales.search_accounts("Industry = 'Technology'")
        assert result == records
