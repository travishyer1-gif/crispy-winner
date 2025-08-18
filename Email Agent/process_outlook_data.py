import json
import os
from typing import Any, Dict, List, Optional

import pandas as pd


def _safe_get_email_name_address(entity: Optional[Dict[str, Any]]) -> (str, str):
	"""Extract name and address from Graph emailAddress dict."""
	if not entity:
		return "", ""
	email_info = entity.get("emailAddress") if "emailAddress" in entity else entity
	name = email_info.get("name", "") if isinstance(email_info, dict) else ""
	address = email_info.get("address", "") if isinstance(email_info, dict) else ""
	return name or "", address or ""


def _extract_sender(item: Dict[str, Any]) -> (str, str):
	# Emails use 'from', events use 'organizer'
	if "from" in item and item["from"]:
		return _safe_get_email_name_address(item.get("from"))
	if "organizer" in item and item["organizer"]:
		return _safe_get_email_name_address(item.get("organizer"))
	return "", ""


def _extract_recipients(item: Dict[str, Any]) -> (str, str):
	# Emails use 'toRecipients' (list); events use 'attendees' (list)
	recipient_list: List[Dict[str, Any]] = []
	if "toRecipients" in item and isinstance(item.get("toRecipients"), list):
		recipient_list = item.get("toRecipients") or []
	elif "attendees" in item and isinstance(item.get("attendees"), list):
		recipient_list = item.get("attendees") or []

	if not recipient_list:
		return "", ""

	names: List[str] = []
	addresses: List[str] = []
	for rec in recipient_list:
		name, addr = _safe_get_email_name_address(rec)
		if name:
			names.append(name)
		if addr:
			addresses.append(addr)

	return "; ".join(names), "; ".join(addresses)


def _extract_date(item: Dict[str, Any]) -> str:
	# Prefer receivedDateTime, then sentDateTime, then event start.dateTime
	for key in ("receivedDateTime", "sentDateTime"):
		if item.get(key):
			return item.get(key)
	# Events
	start = item.get("start")
	if isinstance(start, dict) and start.get("dateTime"):
		return start.get("dateTime")
	return ""


def _extract_body_content(item: Dict[str, Any]) -> str:
	# Use bodyPreview if present; else body.content
	if item.get("bodyPreview"):
		return item.get("bodyPreview") or ""
	body = item.get("body")
	if isinstance(body, dict) and body.get("content"):
		return body.get("content") or ""
	return ""


def _extract_has_attachments(item: Dict[str, Any]) -> bool:
	val = item.get("hasAttachments")
	return bool(val) if isinstance(val, (bool,)) else False


def _extract_attachment_names(item: Dict[str, Any]) -> List[str]:
	attachments = item.get("attachments")
	if isinstance(attachments, list):
		return [att.get("name") for att in attachments if isinstance(att, dict) and att.get("name")]
	return []


def _extract_is_flagged(item: Dict[str, Any]) -> bool:
	# Graph uses flag: { status: notFlagged | complete | flagged }
	flag = item.get("flag")
	if isinstance(flag, dict):
		return (flag.get("status") or "").lower() == "flagged"
	return False


def _first_n_words(text: str, n: int) -> str:
	if not text:
		return ""
	words = str(text).split()
	return " ".join(words[:n])


def normalize_outlook_json(raw: Dict[str, Any]) -> pd.DataFrame:
	"""
	Normalize the combined outlook_data.json into a flat DataFrame.
	Rows represent either emails (inbox/sent) or events.
	"""
	inbox = raw.get("inbox_emails") or []
	sent = raw.get("sent_emails") or []
	events = raw.get("calendar_events") or []

	items: List[Dict[str, Any]] = []

	def append_items(source_list: List[Dict[str, Any]], item_type: str) -> None:
		for it in source_list:
			items.append({"__type": item_type, **it})

	append_items(inbox, "inbox")
	append_items(sent, "sent")
	append_items(events, "event")

	# Build structured rows
	structured: List[Dict[str, Any]] = []
	for item in items:
		item_id = item.get("id", "")
		sender_name, sender_address = _extract_sender(item)
		recipient_name, recipient_address = _extract_recipients(item)
		subject = item.get("subject") or ""
		date_str = _extract_date(item)
		body_content = _extract_body_content(item)
		has_attachment = _extract_has_attachments(item)
		attachment_names = _extract_attachment_names(item)
		is_flagged = _extract_is_flagged(item)

		row: Dict[str, Any] = {
			"id": item_id,
			"record_type": item.get("__type"),
			"sender_name": sender_name,
			"sender_address": sender_address,
			"recipient_name": recipient_name,
			"recipient_address": recipient_address,
			"subject": subject,
			"date": date_str,
			"body_content": body_content,
			"has_attachment": has_attachment,
			"attachment_names": attachment_names,
			"is_flagged": is_flagged,
		}

		# Enrichments
		row["communication_flow"] = (
			f"From: {sender_name or sender_address} To: {recipient_name or recipient_address}".strip()
		)
		row["summary"] = (subject or "").strip()
		body_snippet = _first_n_words(body_content, 50)
		if body_snippet:
			row["summary"] = f"{row['summary']} | {body_snippet}".strip()

		structured.append(row)

	df = pd.DataFrame(structured)

	# Clean data
	if "subject" in df.columns:
		df["subject"] = df["subject"].fillna("")
		df.loc[df["subject"].str.len() == 0, "subject"] = "(no subject)"

	# Deduplicate by id
	if "id" in df.columns:
		df = df.drop_duplicates(subset=["id"])  # keep first occurrence

	# Ensure booleans and lists are proper types
	if "has_attachment" in df.columns:
		df["has_attachment"] = df["has_attachment"].fillna(False).astype(bool)
	if "is_flagged" in df.columns:
		df["is_flagged"] = df["is_flagged"].fillna(False).astype(bool)
	if "attachment_names" in df.columns:
		df["attachment_names"] = df["attachment_names"].apply(lambda v: v if isinstance(v, list) else ([] if pd.isna(v) else [v]))

	return df.reset_index(drop=True)


def main() -> None:
	import argparse
	parser = argparse.ArgumentParser(description="Process Outlook JSON into a flat, enriched table")
	parser.add_argument("--input", "-i", default="outlook_data.json", help="Path to input outlook_data.json")
	parser.add_argument("--output", "-o", default="outlook_data_processed.csv", help="Path to output CSV file")
	parser.add_argument("--output-json", default="outlook_data_processed.json", help="Optional JSON output path")
	args = parser.parse_args()

	if not os.path.exists(args.input):
		raise FileNotFoundError(f"Input file not found: {args.input}")

	with open(args.input, "r", encoding="utf-8") as f:
		raw = json.load(f)

	df = normalize_outlook_json(raw)

	# Save outputs
	df.to_csv(args.output, index=False)
	# Save JSON as records for downstream flexibility
	df.to_json(args.output_json, orient="records", force_ascii=False, indent=2)

	print(f"Processed rows: {len(df)}")
	print(f"Saved CSV: {args.output}")
	print(f"Saved JSON: {args.output_json}")


if __name__ == "__main__":
	main()


