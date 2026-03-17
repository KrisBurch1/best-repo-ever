"""Tests for salesforce/service_cloud.py"""

from unittest.mock import MagicMock, call

import pytest

from salesforce.connection import SalesforceConnection
from salesforce.service_cloud import ServiceCloud


@pytest.fixture()
def mock_conn():
    conn = MagicMock(spec=SalesforceConnection)
    conn.client = MagicMock()
    return conn


@pytest.fixture()
def service(mock_conn):
    return ServiceCloud(mock_conn)


class TestCreateCase:
    def test_required_fields(self, service, mock_conn):
        mock_conn.client.Case.create.return_value = {
            "id": "5001x000001ABC",
            "success": True,
            "errors": [],
        }
        result = service.create_case("Login issue")
        mock_conn.client.Case.create.assert_called_once_with(
            {
                "Subject": "Login issue",
                "Priority": "Medium",
                "Status": "New",
                "Origin": "Web",
            }
        )
        assert result["success"] is True

    def test_optional_fields(self, service, mock_conn):
        mock_conn.client.Case.create.return_value = {"id": "5001x", "success": True, "errors": []}
        service.create_case(
            "Bug",
            description="Crashes on login",
            contact_id="0031x",
            account_id="0011x",
            priority="High",
            status="Working",
            origin="Phone",
        )
        payload = mock_conn.client.Case.create.call_args[0][0]
        assert payload["Description"] == "Crashes on login"
        assert payload["ContactId"] == "0031x"
        assert payload["AccountId"] == "0011x"
        assert payload["Priority"] == "High"
        assert payload["Status"] == "Working"
        assert payload["Origin"] == "Phone"

    def test_extra_kwargs_passed(self, service, mock_conn):
        mock_conn.client.Case.create.return_value = {"id": "5001x", "success": True, "errors": []}
        service.create_case("Subject", Type="Mechanical")
        payload = mock_conn.client.Case.create.call_args[0][0]
        assert payload["Type"] == "Mechanical"


class TestGetCase:
    def test_returns_case(self, service, mock_conn):
        mock_conn.client.Case.get.return_value = {"Id": "5001x", "Subject": "Login issue"}
        result = service.get_case("5001x")
        mock_conn.client.Case.get.assert_called_once_with("5001x")
        assert result["Subject"] == "Login issue"


class TestUpdateCase:
    def test_update_fields(self, service, mock_conn):
        mock_conn.client.Case.update.return_value = 204
        status_code = service.update_case("5001x", Status="Working", Priority="High")
        mock_conn.client.Case.update.assert_called_once_with(
            "5001x", {"Status": "Working", "Priority": "High"}
        )
        assert status_code == 204


class TestCloseCase:
    def test_sets_status_closed(self, service, mock_conn):
        mock_conn.client.Case.update.return_value = 204
        service.close_case("5001x")
        mock_conn.client.Case.update.assert_called_once_with("5001x", {"Status": "Closed"})


class TestSearchCases:
    def test_query_construction(self, service, mock_conn):
        mock_conn.client.query_all.return_value = {"records": []}
        service.search_cases("Status = 'New'")
        mock_conn.client.query_all.assert_called_once_with(
            "SELECT Id, CaseNumber, Subject, Status, Priority FROM Case WHERE Status = 'New'"
        )

    def test_returns_records(self, service, mock_conn):
        records = [{"Id": "5001x", "Subject": "Bug"}]
        mock_conn.client.query_all.return_value = {"records": records}
        result = service.search_cases("Status = 'New'")
        assert result == records


class TestAddCaseComment:
    def test_creates_comment(self, service, mock_conn):
        mock_conn.client.CaseComment.create.return_value = {
            "id": "08a1x",
            "success": True,
            "errors": [],
        }
        result = service.add_case_comment("5001x", "Looking into this now.")
        mock_conn.client.CaseComment.create.assert_called_once_with(
            {"ParentId": "5001x", "CommentBody": "Looking into this now.", "IsPublished": True}
        )
        assert result["success"] is True

    def test_unpublished_comment(self, service, mock_conn):
        mock_conn.client.CaseComment.create.return_value = {"id": "08a1x", "success": True, "errors": []}
        service.add_case_comment("5001x", "Internal note", is_published=False)
        payload = mock_conn.client.CaseComment.create.call_args[0][0]
        assert payload["IsPublished"] is False


class TestCreateContact:
    def test_required_fields(self, service, mock_conn):
        mock_conn.client.Contact.create.return_value = {"id": "0031x", "success": True, "errors": []}
        service.create_contact("Smith")
        mock_conn.client.Contact.create.assert_called_once_with({"LastName": "Smith"})

    def test_optional_fields(self, service, mock_conn):
        mock_conn.client.Contact.create.return_value = {"id": "0031x", "success": True, "errors": []}
        service.create_contact(
            "Smith",
            first_name="John",
            email="john@example.com",
            phone="555-1234",
            account_id="0011x",
        )
        payload = mock_conn.client.Contact.create.call_args[0][0]
        assert payload["FirstName"] == "John"
        assert payload["Email"] == "john@example.com"
        assert payload["Phone"] == "555-1234"
        assert payload["AccountId"] == "0011x"


class TestGetContact:
    def test_returns_contact(self, service, mock_conn):
        mock_conn.client.Contact.get.return_value = {"Id": "0031x", "LastName": "Smith"}
        result = service.get_contact("0031x")
        mock_conn.client.Contact.get.assert_called_once_with("0031x")
        assert result["LastName"] == "Smith"
