"""Dump every sheet of data/revenue-source.xlsx to data/data.json (used by build_deck.js)."""
import json, sys, openpyxl
src, dst = sys.argv[1:3] if len(sys.argv) > 2 else ("data/revenue-source.xlsx", "data/data.json")
wb = openpyxl.load_workbook(src, data_only=True)
out = {ws.title.strip(): [[c.replace("\xa0", " ") if isinstance(c, str) else c for c in r]
                          for r in ws.iter_rows(values_only=True)] for ws in wb}
json.dump(out, open(dst, "w"), ensure_ascii=False)
