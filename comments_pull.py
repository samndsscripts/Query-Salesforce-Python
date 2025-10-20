from simple_salesforce import Salesforce
import re

# --- Salesforce login ---
sf = Salesforce(
    username='samuelcooper@ndspro.com',
    password='Summer@NDS2025',
    security_token='zjU2IJAfQmx6zDxgOj3aLkyPQ',
    instance_url='https://nds.my.salesforce.com'
)

# --- Helper: parse case number safely ---
def parse_case_number(x):
    try:
        return int(float(str(x).strip()))
    except:
        return None

# --- Pull Salesforce report ---
report_id = "00OUI00000EsGR72AN"
report_data = sf.restful(f'analytics/reports/{report_id}', params={'includeDetails': 'true'})
report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', [])

# --- Take only the first 5 cases ---
first_five_cases = []
for row in report_rows[:5]:
    cells = row.get('dataCells', [])
    if len(cells) > 3:
        case_number = cells[3].get('label', '').strip()
        if case_number and re.match(r'^\d+$', case_number):
            first_five_cases.append(case_number)

print(f"🔍 Checking comments for the first {len(first_five_cases)} cases...\n")

# --- Fetch and print case comments ---
for num in first_five_cases:
    # Get Case ID
    case_query = f"SELECT Id FROM Case WHERE CaseNumber = '{num}'"
    case_result = sf.query(case_query)
    case_records = case_result.get('records', [])
    if not case_records:
        print(f"⚠️ Case {num} not found via SOQL.\n")
        continue

    case_id = case_records[0]['Id']

    # Query for comments
    comment_query = f"""
        SELECT CommentBody, CreatedDate, CreatedById
        FROM CaseComment
        WHERE ParentId = '{case_id}'
        ORDER BY CreatedDate DESC
    """
    comments = sf.query(comment_query)
    comment_records = comments.get('records', [])

    print(f"💬 Comments for Case {num}:")
    if comment_records:
        for c in comment_records:
            print(f"   - {c['CreatedDate']}: {c['CommentBody']}\n")
    else:
        print("   (No comments found)\n")
