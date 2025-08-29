from __future__ import annotations
import re
from typing import List

def parse_notes_md(md_text: str, slide_count: int) -> List[str]:
	# Match second-level headers like '## Slide 3'
	header_re = re.compile(r"^##\s+Slide\s+(\d+)\s*$", re.IGNORECASE | re.MULTILINE)
	notes = [""] * slide_count
	matches = list(header_re.finditer(md_text))
	seen_sections = set()
	for i, m in enumerate(matches):
		start = m.end()
		end = matches[i+1].start() if i+1 < len(matches) else len(md_text)
		idx = int(m.group(1)) - 1
		section_text = md_text[start:end].strip()
		if 0 <= idx < slide_count:
			notes[idx] = section_text
			seen_sections.add(idx+1)
	# Warn on unexpected sections: any '## <something>' that doesn't match 'Slide N'
	for hm in re.finditer(r"^##\s+(.*)$", md_text, re.MULTILINE):
		title = hm.group(1).strip()
		if not re.match(r"(?i)^Slide\s+\d+$", title):
			print(f"Warning: unexpected section '## {title}' found in markdown and ignored")
	return notes

def parse_notes(path: str, slide_count: int) -> List[str]:
	if not path.lower().endswith('.md'):
		raise ValueError("Only Markdown (.md) notes are supported in the current configuration.")
	with open(path, 'r', encoding='utf-8') as f:
		text = f.read()
	return parse_notes_md(text, slide_count)

