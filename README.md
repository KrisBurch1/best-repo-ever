# best-repo-ever

## Salesforce Service and Sales Cloud Integration

A Python library for connecting to and working with Salesforce **Service Cloud** and **Sales Cloud** via the [Salesforce REST API](https://developer.salesforce.com/docs/atlas.en-us.api_rest.meta/api_rest/).

---

## Requirements

- Python 3.8+
- A Salesforce org (Developer Edition, Sandbox, or Production)
- User credentials **or** a Connected App with consumer key/secret

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Configuration

Set credentials via environment variables (recommended) or pass them directly to `SalesforceConnection`.

| Variable              | Description                                      | Default   |
|-----------------------|--------------------------------------------------|-----------|
| `SF_USERNAME`         | Salesforce username                              | —         |
| `SF_PASSWORD`         | Salesforce password                              | —         |
| `SF_SECURITY_TOKEN`   | Salesforce security token                        | `""`      |
| `SF_CONSUMER_KEY`     | Connected app consumer key (optional)            | —         |
| `SF_CONSUMER_SECRET`  | Connected app consumer secret (optional)         | —         |
| `SF_DOMAIN`           | Login domain (`login` for prod, `test` for sandbox) | `login` |

---

## Usage

### Connecting

```python
from salesforce import SalesforceConnection, ServiceCloud, SalesCloud

# Reads credentials from environment variables
conn = SalesforceConnection()

# Or pass credentials directly
conn = SalesforceConnection(
    username="user@example.com",
    password="mypassword",
    security_token="mytoken",
    domain="login",          # use "test" for sandboxes
)
```

### Service Cloud — Cases

```python
service = ServiceCloud(conn)

# Create a Case
case = service.create_case(
    subject="Cannot login",
    description="User gets 403 on the login page.",
    priority="High",
    origin="Web",
)
print(case["id"])  # e.g. "5001x000001ABCxABC"

# Retrieve a Case
record = service.get_case(case["id"])

# Add a comment
service.add_case_comment(case["id"], "Assigned to Tier 2 support.")

# Update status
service.update_case(case["id"], Status="Working")

# Close a Case
service.close_case(case["id"])

# Search Cases
open_high = service.search_cases("Status != 'Closed' AND Priority = 'High'")
```

### Service Cloud — Contacts

```python
contact = service.create_contact(
    last_name="Smith",
    first_name="Jane",
    email="jane.smith@example.com",
    phone="555-1234",
)

record = service.get_contact(contact["id"])
```

### Sales Cloud — Leads

```python
sales = SalesCloud(conn)

lead = sales.create_lead(
    last_name="Doe",
    company="Acme Corp",
    first_name="John",
    email="john.doe@acme.com",
    lead_source="Web",
)

sales.update_lead(lead["id"], Status="Working - Contacted")
sales.convert_lead(lead["id"])

new_leads = sales.search_leads("Status = 'Open - Not Contacted'")
```

### Sales Cloud — Opportunities

```python
opp = sales.create_opportunity(
    name="Acme Corp Q2 Deal",
    stage="Prospecting",
    close_date="2026-06-30",
    amount=75000.0,
)

sales.update_opportunity(opp["id"], StageName="Qualification")
sales.close_opportunity_won(opp["id"])

pipeline = sales.search_opportunities("StageName != 'Closed Won' AND StageName != 'Closed Lost'")
```

### Sales Cloud — Accounts

```python
account = sales.create_account(
    name="Acme Corp",
    industry="Technology",
    phone="555-0100",
    website="https://acme.com",
)

sales.update_account(account["id"], Industry="Software")
tech_accounts = sales.search_accounts("Industry = 'Technology'")
```

---

## Running Tests

```bash
pip install -r requirements-dev.txt
pytest tests/ -v
```

---

## Project Structure

```
salesforce/
  __init__.py        # Package exports
  connection.py      # Authentication & connection management
  service_cloud.py   # Service Cloud: Cases, Contacts, Case Comments
  sales_cloud.py     # Sales Cloud: Leads, Opportunities, Accounts
tests/
  test_connection.py
  test_service_cloud.py
  test_sales_cloud.py
requirements.txt
requirements-dev.txt
```
