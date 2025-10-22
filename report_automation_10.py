REPORT_ID = "00OUI00000EsGR72AN"
CYCLE_SECONDS = 60
SOQL_BATCH = 100

try:
    while True:
        cycle_start_time = time.time()  # reset timer each cycle
        clear_console()
        print("🔄 Checking for new Salesforce cases... (Time elapsed: 00:00:00)\n")

        # --- Load existing Excel data ---
        table_vals = table_out.range.value
        if table_vals and len(table_vals) > 1:
            output_df = pd.DataFrame(table_vals[1:], columns=table_vals[0])
        else:
            output_df = pd.DataFrame(columns=table_vals[0] if table_vals else [])

        parsed_cases = output_df.get('Case Number', pd.Series([])).apply(parse_case_number).tolist()
        existing_cases = set(filter(None, parsed_cases))

        # --- Pull Salesforce report ---
        try:
            report_data = sf.restful(f'analytics/reports/{REPORT_ID}', params={'includeDetails': 'true'})
            report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', [])
        except Exception as e:
            print(Fore.RED + "❌ Error fetching report:", str(e))
            time.sleep(CYCLE_SECONDS)
            continue

        # --- Collect new case numbers ---
        new_case_numbers = []
        for row in report_rows:
            cells = row.get('dataCells', [])
            if len(cells) > 3:
                cn_raw = str(cells[3].get('label', '')).strip()
                cn_int = parse_case_number(cn_raw)
                if cn_int and cn_int not in existing_cases:
                    new_case_numbers.append(cn_raw)

        print(f"Found {len(new_case_numbers)} new case(s).")

        if not new_case_numbers:
            print(f"No new cases.\n")

        # --- Batch SOQL: get titles and comments properly ---
        case_subject_map = {}
        case_comments_map = {}
        case_number_to_id = {}  # Map CaseNumber → Case ID

        for batch in [new_case_numbers[i:i+SOQL_BATCH] for i in range(0, len(new_case_numbers), SOQL_BATCH)]:
            quoted = ",".join(f"'{cn}'" for cn in batch)

            # --- Query Case titles (need Ids for next query) ---
            soql_titles = f"""
                SELECT Id, CaseNumber, Subject
                FROM Case
                WHERE CaseNumber IN ({quoted})
            """
            try:
                result_titles = sf.query_all(soql_titles)

                if result_titles is None:
                    print(Fore.YELLOW + f"⚠️ Title query returned None for batch: {batch[:5]}...")
                    continue

                # Collect Case IDs for comment query
                case_ids = []
                for rec in result_titles.get('records', []):
                    case_number = rec.get('CaseNumber', '')
                    subject = rec.get('Subject', None)
                    case_subject_map[case_number] = subject or ''
                    case_number_to_id[case_number] = rec.get('Id', '')
                    if rec.get('Id'):
                        case_ids.append(rec['Id'])

                # --- Query Comments using Case IDs ---
                if case_ids:
                    quoted_ids = ",".join(f"'{cid}'" for cid in case_ids)
                    soql_comments = f"""
                        SELECT ParentId, CommentBody
                        FROM CaseComment
                        WHERE ParentId IN ({quoted_ids})
                        ORDER BY CreatedDate DESC
                    """
                    try:
                        result_comments = sf.query_all(soql_comments)
                        if result_comments and 'records' in result_comments:
                            # Store latest comment per Case ID
                            for rec in result_comments['records']:
                                parent_id = rec.get('ParentId', '')
                                comment_body = rec.get('CommentBody', '')
                                if parent_id not in case_comments_map:
                                    case_comments_map[parent_id] = comment_body
                        else:
                            print(Fore.CYAN + f"⚠️ No comments found for batch {batch[:5]}...")

                    except Exception as e:
                        print(Fore.YELLOW + f"⚠️ Comment query failed: {e}")
                else:
                    print(Fore.CYAN + f"⚠️ No Case IDs found for batch {batch[:5]}...")

            except Exception as e:
                print(Fore.YELLOW + f"⚠️ SOQL title query failed: {e}")

        # --- Build new rows ---
        new_rows_to_add = []
        for row in report_rows:
            cells = row.get('dataCells', [])
            if len(cells) < 10:
                continue
            cn_raw = str(cells[3].get('label', '')).strip()
            cn_int = parse_case_number(cn_raw)
            if not cn_int or cn_int in existing_cases:
                continue
            if cn_raw not in new_case_numbers:
                continue

            description = str(cells[4].get('label', ''))
            case_id = case_number_to_id.get(cn_raw, None)
            comments = case_comments_map.get(case_id, '') if case_id else ''
            subject = case_subject_map.get(cn_raw, '')
            qty = extract_quantity(subject)
            rma_value = cells[5].get('label', '')
            case_category = cells[6].get('label', '')
            account_name = cells[7].get('label', '')
            contact_type = cells[8].get('label', '')
            shipping_whse = cells[9].get('label', '')

            # Determine top source matches
            source, top_matches = determine_source(description, top_n=3)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Add structured match columns
            top1 = top_matches[0] if len(top_matches) > 0 else ''
            top2 = top_matches[1] if len(top_matches) > 1 else ''
            top3 = top_matches[2] if len(top_matches) > 2 else ''

            new_row = [
                cells[0].get('label', ''),  # Opened Date
                cells[1].get('label', ''),  # Case Reason
                cells[2].get('label', ''),  # Case Owner
                cn_int,                     # Case Number
                description,                # Description
                comments,                   # Comments
                qty if qty else '',         # Quantity
                rma_value,                  # RMA Value
                case_category,              # Case Category
                account_name,               # Account Name
                contact_type,               # Contact Type
                shipping_whse,              # Shipping Whse
                top1,                       # Top Match 1
                top2,                       # Top Match 2
                top3,                       # Top Match 3
                source,                     # Source
                timestamp                   # Time Stamp
            ]
            new_rows_to_add.append(new_row)
            existing_cases.add(cn_int)

        # --- Append to Excel table properly ---
        if new_rows_to_add:
            if table_out.data_body_range:
                start_row = table_out.data_body_range.last_cell.row + 1
            else:
                start_row = table_out.range.row + 1

            start_col = table_out.range.column
            sheet_out.range((start_row, start_col)).value = new_rows_to_add

            print(Fore.GREEN + f"\n✅ Added {len(new_rows_to_add)} new row(s) to Excel.\n")
            for nr in new_rows_to_add:
                print(f" → Case {nr[3]} | Qty: {nr[6] or 'N/A'} | Source: {nr[15]} | Added: {nr[16]}")
        else:
            print(Fore.YELLOW + "No valid new rows to add this cycle.")

        # --- Sleep with real-time elapsed display ---
        for i in range(CYCLE_SECONDS):
            elapsed_seconds = int(time.time() - cycle_start_time)
            hours, remainder = divmod(elapsed_seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            elapsed_str = f"{hours:02}:{minutes:02}:{seconds:02}"

            clear_console()
            print(f"🔄 Checking for new Salesforce cases... (Time elapsed: {elapsed_str})\n")
            print(f"Found {len(new_case_numbers)} new case(s).")
            if not new_case_numbers:
                print(f"No new cases.")
            time.sleep(1)

except KeyboardInterrupt:
    print("\n" + Style.BRIGHT + "Script stopped by user.")
