"""Salesforce Service and Sales Cloud connection package."""

from .connection import SalesforceConnection
from .service_cloud import ServiceCloud
from .sales_cloud import SalesCloud

__all__ = ["SalesforceConnection", "ServiceCloud", "SalesCloud"]
