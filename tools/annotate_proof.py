"""Annotate the Frontiers author's proof with the authors' query responses and corrections.

Usage:
    python tools/annotate_proof.py <proof.pdf> <Author_Queries_Response.docx> <out.pdf> <checklist.md>

* Each query response is read verbatim from the Word file. Only the edits listed in
  RESPONSE_EDITS are applied, and every edit must match exactly once.
* Each correction is marked at its place in the proof (strike-out = delete/replace,
  caret = insert, highlight = see comment). Every comment names its query and page/line.
* Each response goes into the "Author's response" column of the queries form.
"""
import re
import shutil
import sys

import docx
import pymupdf

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from proofmap import line_of, margin_numbers, column_side  # noqa: E402

AUTHOR = "Authors"
RED, BLUE, YELLOW = (0.85, 0, 0), (0, 0.2, 0.85), (1, 0.9, 0.2)

# --------------------------------------------------------------------------------------
# Response texts: verbatim from the Word file + minimal edits (old, new). Each must match once.
# --------------------------------------------------------------------------------------
RESPONSE_EDITS = {
    "Q1": [("No changes are needed.", "No changes are needed. [Proof: p.5, l.64-68]")],
    "Q3": [("No changes are needed.", "No changes are needed. [Proof: p.5, l.71 and l.25-28]")],
    "Q4": [("Lixing Liu, Wencai Zhang, Chaobo Zhou.", "Lixing Liu, Wencai Zhang, Chaobo Zhou. [Proof: p.5, l.71]")],
    "Q5": [("No department name is required.", "No department name is required. [Proof: p.5, l.72-73 - no change]")],
    "Q6": [("No changes are needed.", "No changes are needed. [Proof: p.5, l.72-74]")],
    "Q7": [("no changes are needed.", "no changes are needed. [Proof: p.5, l.18-20]")],
    "Q8": [("Please replace the current abstract with",
            "Please replace the current abstract (p.5, l.76-91; marked in the proof) with")],
    "Q9": [("No changes are needed.", "No changes are needed. [Proof: p.5, l.94]")],
    # 6.2.x-6.5.x are also third-level headings; the original sentence listed only 4.1.x-4.3.x.
    "Q10": [("2.x-6.x (secondary) and 4.1.x-4.3.x (tertiary)",
             "2.x-6.x (secondary) and 4.1.x-4.3.x and 6.2.x-6.5.x (tertiary)")],
    "Q11": [
        ("1) Sun and Zeng, 2023 - please ADD the reference entry:",
         "1) Sun and Zeng, 2023 - please ADD the reference entry (before Sun and Zheng, 2024; p.28, l.2752):"),
        ("(Sections 1 and 2.2) are already present and correct.",
         "(Sections 1 and 2.2; p.6, l.122 and p.7, l.287-288) are already present and correct."),
        ("  - Section 1: \"green bonds (Zhang and Guo, 2025; Lau, 2025)\"",
         "  - Section 1 (p.6, l.123): \"green bonds (Zhang and Guo, 2025; Lau, 2025)\""),
        ("  - Section 2.2: \"projects (Zhang and Guo, 2025; Lau, 2025)\"",
         "  - Section 2.2 (p.7, l.293): \"projects (Zhang and Guo, 2025; Lau, 2025)\""),
        ("3) Wang and Jin, 2025 - please ADD the reference entry:",
         "3) Wang and Jin, 2025 - please ADD the reference entry (before Whited and Wu, 2006; p.28, l.2771):"),
        ("  - Section 1: \"...bridging a critical gap",
         "  - Section 1 (p.6, l.197): \"...bridging a critical gap"),
        ("  - Section 2.4: \"...bridges the gaps",
         "  - Section 2.4 (p.8, l.413): \"...bridges the gaps"),
        ("Please correct the reference entry to: Farooq",
         "Please correct the reference entry (p.28, l.2761-2762) to: Farooq"),
        ("Please replace the reference entry with: Xu, Y.",
         "Please replace the reference entry (p.27, l.2601-2602) with: Xu, Y."),
        ("6) Dahiya and Singh, 2021 - please ADD the reference entry:",
         "6) Dahiya and Singh, 2021 - please ADD the reference entry (before DesJardine et al., 2021; p.28, l.2671):"),
        ("Please replace the reference entry with: Cheng, W.",
         "Please replace the reference entry (p.27, l.2599-2600) with: Cheng, W."),
        ("8) China Energy Statistical Yearbook (2022) - please ADD the reference entry:",
         "8) China Energy Statistical Yearbook (2022) - please ADD the reference entry (before Orazalin et al., 2023; p.28, l.2739):"),
        ("  - Section 4.1.2: \"Following Maetal (2023)",
         "  - Section 4.1.2 (p.10, l.670): \"Following Maetal (2023)"),
    ],
    "Q12": [("Section 6.2.4 now cites Equation 5 explicitly:",
             "Section 6.2.4 (p.16, l.1367) now cites Equation 5 explicitly:")],
    "Q13": [("No changes are needed.", "No changes are needed.")],
    "Q14": [
        # The caption text was not given, and the diagnostics sentence pointed at Table 4 instead of Table 5.
        ("- Table 5: the sentence \"TABLE 5 reports the 2SLS estimation results alongside relevant diagnostic tests.\" will be moved into the body text of Section 6.3.1 as \"Table 5 reports the 2SLS estimation results alongside relevant diagnostic tests.\", and a proper caption will be added to the table.",
         "- Table 5 (caption at p.18, l.1509): the sentence \"TABLE 5 reports the 2SLS estimation results alongside relevant diagnostic tests.\" will be moved into the body text of Section 6.3.1 (p.17, l.1500) as \"Table 5 reports the 2SLS estimation results alongside relevant diagnostic tests.\", and a proper caption will be added to the table: \"TABLE 5 Instrumental variable (2SLS) estimation results.\" In the same paragraph (p.17, l.1500), \"The diagnostic test statistics presented in Table 4\" should read \"Table 5\" (Table 4 is the greenwashing robustness test)."),
        ("- Table 7: Section 6.4.1,", "- Table 7: Section 6.4.1 (p.19, l.1645),"),
        ("- Table 9: Section 6.4.3,", "- Table 9: Section 6.4.3 (p.19, l.1733),"),
        ("- Table 10: Section 6.4.4,", "- Table 10: Section 6.4.4 (p.23, l.2143),"),
    ],
    "Q16": [
        ("- Table 2 (row of", "- Table 2 (p.14, l.1105; row of"),
        ("- Table 3 (columns 1-6):", "- Table 3 (p.15, l.1163 and l.1221; columns 1-6):"),
        ("- Table 7 (columns 1-2):", "- Table 7 (p.20, l.1743/1801; columns 1-2):"),
        ("- Table 8 (columns 1-2):", "- Table 8 (p.21, l.1859/1917; columns 1-2):"),
        ("- Table 9 (columns 1-3):", "- Table 9 (p.22, l.1975/2033; columns 1-3):"),
        ("- Table 10 (columns 1-2):", "- Table 10 (p.23, l.2091/2149; columns 1-2):"),
        ("- Table 12 (columns 1-3):", "- Table 12 (p.25, l.2323; columns 1-3):"),
    ],
    "Q18": [("Section 6.4.3, \"Regression", "Section 6.4.3 (p.19, l.1733), \"Regression"),
            ("Section 6.4.4, \"Columns (1)", "Section 6.4.4 (p.23, l.2143), \"Columns (1)")],
    "Q19": [("without undue reservation.\"", "without undue reservation.\" [Proof: p.27, l.2587-2588]")],
    "Q20": [("Please update the Author Contributions section to read:",
             "Please update the Author Contributions section (p.27, l.2613) to read:")],
    "Q21": [("No commercial funding was involved.", "No commercial funding was involved. [Proof: p.27, l.2620-2621]")],
    "Q22": [("No changes are needed.", "No changes are needed. [Proof: p.27, l.2626-2628]")],
    "Q23": [("(i) delete the reference entry;", "(i) delete the reference entry (p.28, l.2741-2743);"),
            ("amend the in-text citation in Section 1,", "amend the in-text citation in Section 1 (p.6, l.145-146),")],
    "Q24": [
        ("- \"Anonymous, 2024\" (= Cheng et al., 2024):", "- \"Anonymous, 2024\" (= Cheng et al., 2024; p.27, l.2599-2600):"),
        ("- \"Anonymous, 2025\" (= Xu and Shi, 2025):", "- \"Anonymous, 2025\" (= Xu and Shi, 2025; p.27, l.2601-2602):"),
        ("- Bratten et al., 2026:", "- Bratten et al., 2026 (p.27, l.2606-2607):"),
        ("- Kavadis and Thomsen, 2022:", "- Kavadis and Thomsen, 2022 (p.28, l.2704-2705):"),
        ("- Li et al., 2024:", "- Li et al., 2024 (p.28, l.2717-2718):"),
        ("- Lian, 2016:", "- Lian, 2016 (p.28, l.2721-2722):"),
        ("- Lu, 2025:", "- Lu, 2025 (p.28, l.2734):"),
        ("- Peng et al., 2024:", "- Peng et al., 2024 (p.28, l.2741-2743):"),
        ("- Sun and Zheng, 2024:", "- Sun and Zheng, 2024 (p.28, l.2752-2753):"),
        ("- \"Umar et al., 2024\" (= Farooq et al., 2024):", "- \"Umar et al., 2024\" (= Farooq et al., 2024; p.28, l.2761-2762):"),
    ],
    "Q25": [("- \"Anonymous, 2025\" (= Xu and Shi, 2025):", "- \"Anonymous, 2025\" (= Xu and Shi, 2025; p.27, l.2601-2602):"),
            ("- Bratten et al., 2026:", "- Bratten et al., 2026 (p.27, l.2606-2607):")],
    "Q26": [("- Holladay, J. S. (2010). Are Exporters Mother Nature's Best Friends? FREIT Working Paper. This",
             "- Holladay, J. S. (2010) (p.28, l.2692-2693). Are Exporters Mother Nature's Best Friends? FREIT Working Paper. This"),
            ("- Lerner, J. (2020).", "- Lerner, J. (2020) (p.28, l.2715-2716).")],
    "Q27": [("- Li et al., 2024:", "- Li et al., 2024 (p.28, l.2717-2718):"),
            ("- Lian, 2016:", "- Lian, 2016 (p.28, l.2721-2722):"),
            ("- Lu, 2025:", "- Lu, 2025 (p.28, l.2734):")],
    "Q28": [("This reference has been retracted", "This reference (p.28, l.2741-2743) has been retracted")],
    "Q29": [
        ("1. Acemoglu, 2007 - Section 3.3:", "1. Acemoglu, 2007 - Section 3.3 (p.9, l.492):"),
        # The quoted sentence did not match the proof; the citation is simply added next to Jin et al. (2017).
        ("2. Cheng et al., 2021 - Section 2.1: added alongside Jin et al. (2017), \"...suppress carbon emission growth (Jin et al., 2017), while technological innovation mitigates CO2 emissions (Cheng et al., 2021)\".",
         "2. Cheng et al., 2021 - Section 2.1 (p.7, l.259): added alongside Jin et al. (2017), \"...suppress carbon emission growth (Jin et al., 2017; Cheng et al., 2021)\"."),
        ("3. Hadlock and Pierce, 2010 - Section 2.3 (first channel, financing constraints):",
         "3. Hadlock and Pierce, 2010 - Section 2.3 (p.7, l.321; first channel, financing constraints):"),
        ("4. Lau et al., 2025 - the in-text \"Lau, 2025\" (Sections 1 and 2.2)",
         "4. Lau et al., 2025 - the in-text \"Lau, 2025\" (Sections 1 and 2.2; p.6, l.123 and p.7, l.293)"),
        ("5. Liu et al., 2022 - Section 2.1:", "5. Liu et al., 2022 - Section 2.1 (p.7, l.261):"),
        ("6. Liu et al., 2024 - Section 2.3:", "6. Liu et al., 2024 - Section 2.3 (p.8, l.353):"),
        # The proof's cluster also contains Lu, 2025, which stays.
        ("7. Roberts and Dowling, 2002 - Section 6.5.2 (corporate reputation mechanism): added to the citation cluster (Dyck et al., 2019; Bolton and Kacperczyk, 2021; Flammer, 2021; Roberts and Dowling, 2002).",
         "7. Roberts and Dowling, 2002 - Section 6.5.2 (p.25, l.2426; corporate reputation mechanism): added to the citation cluster (Dyck et al., 2019; Bolton and Kacperczyk, 2021; Flammer, 2021; Lu, 2025; Roberts and Dowling, 2002)."),
        ("8. Sun and Zheng, 2024 - Section 6.4.2 (regional environmental regulation heterogeneity): added to the discussion of environmental regulation effects.",
         "8. Sun and Zheng, 2024 - Section 6.4.2 (p.19, l.1713; regional environmental regulation heterogeneity): added to the discussion of environmental regulation effects: \"...to adhere to sustainability mandates (Sun and Zheng, 2024).\""),
        # There is no Tan et al. citation in Section 2.2 of the proof.
        ("The in-text citations in Section 1 and Section 2.2 refer to 2025a; those in Sections 2.3 and 3.1 refer to 2025b.",
         "The in-text citation in Section 1 (p.6, l.122) refers to 2025a; those in Sections 2.3 and 3.1 (p.7, l.334 and p.8, l.434) refer to 2025b. Reference list: p.28, l.2754 (2025a) and l.2757 (2025b)."),
        ("10. Tian and Lin, 2019 - Section 2.3 (financing constraints): added alongside Zhang et al. (2019).",
         "10. Tian and Lin, 2019 - Section 2.3 (p.7, l.321; financing constraints): added alongside Zhang et al. (2019)."),
        ("12. Whited and Wu, 2006 - Section 2.3 (measurement of financing constraints): added.",
         "12. Whited and Wu, 2006 - Section 2.3 (p.7, l.321; measurement of financing constraints): added."),
        ("13. Xu et al., 2021 - Section 2.1 (green innovation and carbon intensity): added.",
         "13. Xu et al., 2021 - Section 2.1 (p.7, l.261; green innovation and carbon intensity): added."),
        ("14. Ye and Tian, 2025 - Section 6.5.1 (ESG responsibility mechanism): added to the citation cluster.",
         "14. Ye and Tian, 2025 - Section 6.5.1 (p.24, l.2277; ESG responsibility mechanism): added to the citation cluster (Qian and Liu, 2024; Xie, 2024; Ye and Tian, 2025)."),
        ("15. Zhang et al., 2019 - Section 2.3 (financing constraints): added.",
         "15. Zhang et al., 2019 - Section 2.3 (p.7, l.321; financing constraints): added."),
        ("16. Zhao et al., 2024a - Section 2.2 (green finance and corporate carbon emissions): added.",
         "16. Zhao et al., 2024a - Section 2.2 (p.7, l.283; green finance and corporate carbon emissions): added. Reference list: Zhao, X. et al. (2024) -> 2024a (p.29, l.2844)."),
        ("17. Zhao et al., 2024b - Section 6.6 (moderating mechanisms, digital economy): added.",
         "17. Zhao et al., 2024b - Section 6.6 (p.26, l.2446; moderating mechanisms, digital economy): added. Reference list: Zhao, Z. et al. (2024) -> 2024b (p.29, l.2847)."),
        ("18. Zhou et al., 2025 - Section 2.2 (green finance pilot policy): added.",
         "18. Zhou et al., 2025 - Section 2.2 (p.7, l.283; green finance pilot policy): added."),
        ("21. Xue and Ai, 2025 - Section 2.3 (supply chain finance and carbon emission reduction): added.",
         "21. Xue and Ai, 2025 - Section 2.3 (p.7, l.334; supply chain finance and carbon emission reduction): added."),
        ("(\"emphasized by Li et al.\", \"e.g., Li et al.\")",
         "(\"emphasized by Li et al.\", p.6, l.185; \"e.g., Li et al.\", p.8, l.413)"),
    ],
}

# Additional corrections not raised by the typesetter (renumbered E1-E7 for submission).
ADDITIONAL = [
    ("E1", "Section 6.2.3 (p.16, l.1346-1355) contains a duplicated paragraph (\"Certain external shocks ... carbon emissions performance.\" appears twice). Please delete the duplicate so the paragraph appears only once."),
    ("E2", "Table 2 note (p.14, l.1150) and Table 4 note (p.17, l.1407): \"*, **, *** denote significance at the 1%, 5%, and 10% levels respectively\" is inverted. Please change to \"***, **, * denote significance at the 1%, 5%, and 10% levels respectively\"."),
    ("E3", "Section 6.2.1 (p.16, l.1310-1311): \"In Column Column (2) adds province x year fixed effects\" -> \"Column (2) adds province x year fixed effects\"."),
    ("E4", "Section 3.4 (p.9, l.541): \"establishing the baseline baseline effect\" -> \"establishing the baseline effect\"."),
    ("E5", "Section 3.3 (p.9, l.503): \"while a individual CEO's awareness\" -> \"while an individual CEO's awareness\"."),
    ("E6", "Section 4.1 (p.10, l.588): \"the primary explanatory variable of interest., The\" -> remove the extra comma (\"of interest. The\")."),
    ("E7", "Section 6.2.4, description of Equation 5 (p.16, l.1368 and l.1371-1372): the definitions of Costit and Costjt are transposed - \"Costit captures the operating costs of firm j in year t, whereas Costjt indicates the aggregate operating costs of industry i\". Please correct to \"Costjt captures the operating costs of firm j in year t, whereas Costit indicates the aggregate operating costs of industry i\", as in the equation. Also add the missing full stop after \"in year t\" (l.1368): \"...of firm j in year t. NCit refers to...\"."),
]

# Proof locations for responses that need no change in the article.
EXTRA_NOTE = {
    "Q2": None, "Q13": None, "Q15": None, "Q17": None, "Q30": None,
}


def load_responses(path):
    table = docx.Document(path).tables[0]
    out = {}
    for row in table.rows[1:]:
        q, text = row.cells[0].text.strip(), row.cells[1].text.strip()
        out[q] = text
    for q, edits in RESPONSE_EDITS.items():
        for old, new in edits:
            n = out[q].count(old)
            if n != 1:
                raise SystemExit(f"{q}: edit target found {n} times: {old[:60]!r}")
            out[q] = out[q].replace(old, new)
    return out


# --------------------------------------------------------------------------------------
# Locating text in the proof
# --------------------------------------------------------------------------------------
class Proof:
    def __init__(self, path):
        self.doc = pymupdf.open(path)
        self.log = []  # (queries, page, lines, kind, instruction)
        self._pages = {}

    def page(self, pno):
        # Keep page objects alive: annotations are bound to the Page they were created on.
        if pno not in self._pages:
            self._pages[pno] = self.doc[pno - 1]
        return self._pages[pno]

    def find(self, pno, text, line, nth=0):
        page = self.page(pno)
        nums = margin_numbers(page)
        hits = sorted((r for r in page.search_for(text) if line_of(page, r, nums) == line), key=lambda r: r.x0)
        # A phrase can come back in pieces (one per text span); join pieces that touch.
        merged = []
        for r in hits:
            if merged and 0 <= r.x0 - merged[-1].x1 < 15 and text.count(" ") > 0 and \
                    not self._complete(page, merged[-1], text):
                merged[-1] = merged[-1] | r
            else:
                merged.append(pymupdf.Rect(r))
        hits = merged
        if len(hits) <= nth:
            raise SystemExit(f"p.{pno} l.{line}: {text!r} not found")
        return hits[nth]

    @staticmethod
    def _complete(page, rect, text):
        """True when the rectangle already holds the whole phrase."""
        got = page.get_textbox(rect + (-0.5, -0.5, 0.5, 0.5))
        norm = lambda t: re.sub(r"\s+", "", t).lower()
        return norm(text) in norm(got)

    def line_span(self, pno, side, line, x_from=None, x_to=None):
        page = self.page(pno)
        nums = margin_numbers(page)
        rects = []
        for x0, y0, x1, y1, txt, *_ in page.get_text("words"):
            if x1 < 45 or x0 > page.rect.width - 45:
                continue
            r = pymupdf.Rect(x0, y0, x1, y1)
            if column_side(page, r) != side or line_of(page, r, nums) != line:
                continue
            if x_from is not None and x1 <= x_from + 0.5:
                continue
            if x_to is not None and x0 >= x_to - 0.5:
                continue
            rects.append(r)
        if not rects:
            raise SystemExit(f"p.{pno} l.{line}: empty line")
        u = rects[0]
        for r in rects[1:]:
            u |= r
        if x_from is not None:
            u.x0 = max(u.x0, x_from)
        if x_to is not None:
            u.x1 = min(u.x1, x_to)
        return u

    def block(self, pno, side, first, last, start=None, end=None):
        """Rects covering lines first..last; start/end = (text, line) trims the first/last line."""
        rects = []
        for ln in range(first, last + 1):
            x_from = self.find(pno, *start).x0 if (start and ln == first) else None
            x_to = self.find(pno, *end).x1 if (end and ln == last) else None
            try:
                rects.append(self.line_span(pno, side, ln, x_from, x_to))
            except SystemExit:
                continue  # blank line inside the block
        return rects

    # ---- annotation primitives ----
    def _finish(self, annot, queries, content, color):
        annot.set_colors(stroke=color)
        annot.set_info(title=AUTHOR, content=content, subject=", ".join(queries))
        annot.update()

    def strike(self, queries, pno, rects, lines, instruction):
        content = f"[{', '.join(queries)}] p.{pno}, l.{lines}: {instruction}"
        a = self.page(pno).add_strikeout_annot(rects if isinstance(rects, list) else [rects])
        self._finish(a, queries, content, RED)
        self.log.append((queries, pno, lines, "Strike-out", instruction))

    def highlight(self, queries, pno, rects, lines, instruction):
        content = f"[{', '.join(queries)}] p.{pno}, l.{lines}: {instruction}"
        a = self.page(pno).add_highlight_annot(rects if isinstance(rects, list) else [rects])
        self._finish(a, queries, content, YELLOW)
        self.log.append((queries, pno, lines, "Highlight", instruction))

    def caret(self, queries, pno, rect, lines, instruction, before=False):
        content = f"[{', '.join(queries)}] p.{pno}, l.{lines}: {instruction}"
        x = rect.x0 if before else rect.x1
        page = self.page(pno)
        a = page.add_caret_annot(pymupdf.Point(x - 4, rect.y1 - 4))
        a.set_rect(pymupdf.Rect(x - 4, rect.y1 - 5, x + 4, rect.y1 + 2))
        self._finish(a, queries, content, BLUE)
        self.log.append((queries, pno, lines, "Insert (caret)", instruction))


# --------------------------------------------------------------------------------------
# In-text corrections
# --------------------------------------------------------------------------------------
STRUCTURED_ABSTRACT_NOTE = "Replace the whole abstract with the structured (IMRaD) abstract given in the response to Q8 (Background / Objective / Methods / Results / Conclusion)."


def mark_corrections(p, responses):
    f = p.find
    # Q8 - abstract
    p.strike(["Q8"], 5, p.block(5, "R", 76, 91), "76-91",
             STRUCTURED_ABSTRACT_NOTE + "\n\n" + responses["Q8"].split("abstract:", 1)[1].strip())

    # Q11 / Q29 - in-text citations
    p.strike(["Q11", "Q29"], 6, f(6, "Lau, 2025", 123), "123", 'Replace "Lau, 2025" with "Lau et al., 2025".')
    p.strike(["Q11", "Q29"], 7, f(7, "Lau, 2025", 293), "293", 'Replace "Lau, 2025" with "Lau et al., 2025".')
    p.caret(["Q11", "Q29"], 6, f(6, "Wang and Jin", 197), "197",
            'Insert " (2025)" after "Wang and Jin": "...in the research by Wang and Jin (2025)."')
    p.strike(["Q11", "Q29"], 8, f(8, "(e.g., Li et al.; Wang and Jin)", 413), "413",
             'Replace "(e.g., Li et al.; Wang and Jin)" with "(e.g., Li et al., 2025; Wang and Jin, 2025)".')
    p.strike(["Q29"], 6, f(6, "Li et al.", 185), "185",
             'Replace "Li et al." with "Li et al. (2025)." - "...perspectives emphasized by Li et al. (2025). While previous literature..."')
    p.strike(["Q11"], 10, f(10, "Maetal (2023)", 670), "670", 'Replace "Maetal (2023)" with "Chi et al. (2023)".')

    # Q29 - missing citations
    p.caret(["Q29"], 9, f(9, "Acemoglu et al.", 492), "492",
            'Insert "Acemoglu (2007) and " before "Acemoglu et al. (2012)": "Acemoglu (2007) and Acemoglu et al. (2012) note that..."', before=True)
    p.caret(["Q29"], 7, f(7, "Jin et al., 2017", 259), "259", 'Insert "; Cheng et al., 2021": "(Jin et al., 2017; Cheng et al., 2021)".')
    p.caret(["Q29"], 7, f(7, "Moaniba, 2017", 261), "261",
            'Insert "; Liu et al., 2022; Xu et al., 2021": "(Su and Moaniba, 2017; Liu et al., 2022; Xu et al., 2021)".')
    p.caret(["Q29"], 7, f(7, "Chi et al., 2023", 321), "321",
            'Insert "; Hadlock and Pierce, 2010; Whited and Wu, 2006; Tian and Lin, 2019; Zhang et al., 2019": '
            '"(Guo and Fang, 2024; Shao et al., 2023; Chi et al., 2023; Hadlock and Pierce, 2010; Whited and Wu, 2006; Tian and Lin, 2019; Zhang et al., 2019)".')
    p.caret(["Q29"], 7, f(7, "reduce carbon emissions", 283), "283",
            'Insert " (Zhao et al., 2024a; Zhou et al., 2025)" before the full stop: "...and reduce carbon emissions (Zhao et al., 2024a; Zhou et al., 2025)."')
    p.caret(["Q29"], 7, f(7, "et al., 2025", 334), "334",
            'Change "Tan et al., 2025" to "Tan et al., 2025b" and add "; Xue and Ai, 2025": "(Li et al., 2024; Guo et al., 2024; Tan et al., 2025b; Xue and Ai, 2025)".')
    p.caret(["Q29"], 6, f(6, "Tan et al., 2025", 122), "122", 'Change "Tan et al., 2025" to "Tan et al., 2025a".')
    p.caret(["Q29"], 8, f(8, "et al., 2025", 434), "434", 'Change "Tan et al., 2025" to "Tan et al., 2025b".')
    p.caret(["Q29"], 8, f(8, "Liu et al.", 353), "353",
            'Insert " (2024)" after "Liu et al.": "Research by Liu et al. (2024) and Wang et al. (2024) indicates..."')
    p.caret(["Q29"], 19, f(19, "mandates", 1713), "1713",
            'Insert " (Sun and Zheng, 2024)" before the full stop: "...to adhere to sustainability mandates (Sun and Zheng, 2024)."')
    p.caret(["Q29"], 24, f(24, "Xie, 2024", 2277), "2277", 'Insert "; Ye and Tian, 2025": "(Qian and Liu, 2024; Xie, 2024; Ye and Tian, 2025)".')
    p.caret(["Q29"], 25, f(25, "Lu, 2025", 2426), "2426",
            'Insert "; Roberts and Dowling, 2002": "(Dyck et al., 2019; Bolton and Kacperczyk, 2021; Flammer, 2021; Lu, 2025; Roberts and Dowling, 2002)".')
    p.caret(["Q29"], 26, f(26, "capabilities", 2446), "2446",
            'Insert " (Zhao et al., 2024b)" before the full stop: "...possessing advanced digital technology capabilities (Zhao et al., 2024b)."')

    # Q12 - Equation 5 citation
    p.strike(["Q12"], 16, f(16, "In the equation above", 1367), "1367", 'Replace "In the equation above" with "In Equation 5".')

    # Q14 / Q18 - table citations
    p.highlight(["Q14"], 18, f(18, "TABLE 5 reports the 2SLS estimation results alongside relevant diagnostic tests.", 1509), "1509",
                'This sentence is not a caption. Move it into the text of Section 6.3.1 (see p.17, l.1500) as "Table 5 reports the 2SLS estimation results alongside relevant diagnostic tests." and use the caption "TABLE 5 Instrumental variable (2SLS) estimation results."')
    p.caret(["Q14"], 17, f(17, "The diagnostic test statistics", 1500), "1500",
            'Insert the sentence moved from the Table 5 caption: "Table 5 reports the 2SLS estimation results alongside relevant diagnostic tests. "', before=True)
    p.strike(["Q14"], 17, f(17, "Table 4", 1500), "1500",
             'Replace "Table 4" with "Table 5" (the IV diagnostic statistics are reported in Table 5; Table 4 is the greenwashing test).')
    p.strike(["Q14"], 19, f(19, "Table 6", 1645), "1645", 'Replace "Table 6" with "Table 7".')
    p.strike(["Q14", "Q18"], 19, f(19, "Table 14", 1733), "1733", 'Replace "Table 14" with "Table 9".')
    p.strike(["Q14", "Q18"], 23, f(23, "Table 15", 2143), "2143", 'Replace "Table 15" with "Table 10".')

    # Q16 - column headings
    heads = {
        2: (14, [("(1)", 1105), ("(2)", 1105), ("(3)", 1105), ("(4)", 1105)], "1105",
            "(1) Without control variables; (2) With control variables (pooled OLS); (3) With control variables and firm fixed effects; (4) With control variables, firm fixed effects and industry x year fixed effects."),
        3: (15, [("(1)", 1163), ("(2)", 1163), ("(3)", 1163), ("(4)", 1221), ("(5)", 1221), ("(6)", 1221)], "1163, 1221",
            '(1) Industry and year fixed effects; (2) Firm + industry x year + province x year fixed effects; (3) Standard errors clustered at industry x province level; (4) Excluding observations from 2020 onwards; (5) Alternative dependent variable (CE); (6) 5% tail trimming. (The existing group heading "Modify fixed effects" may be kept as a spanning header over columns (1)-(2).)'),
        7: (20, [("(1)", 1743), ("(2)", 1801)], "1743, 1801",
            "(1) Non-heavily polluting industries; (2) Heavily polluting industries."),
        8: (21, [("(1)", 1859), ("(2)", 1917)], "1859, 1917",
            "(1) Weak environmental regulation; (2) Stringent environmental regulation."),
        9: (22, [("(1)", 1975), ("(2)", 2033), ("(3)", 2033)], "1975, 2033",
            "(1) Growth phase; (2) Maturity stage; (3) Decline phase."),
        10: (23, [("(1)", 2091), ("(2)", 2149)], "2091, 2149",
             "(1) Low industry concentration; (2) High industry concentration."),
        12: (25, [("(1)", 2323), ("(2)", 2323), ("(3)", 2323)], "2323",
             "(1) Digital technology moderation; (2) Managers' education moderation; (3) Artificial intelligence moderation."),
    }
    for t, (pno, parts, lines, text) in heads.items():
        p.highlight(["Q16"], pno, [f(pno, s, ln) for s, ln in parts], lines,
                    f"Table {t} column headings: {text} Keep the existing second header row (dependent variable) as the sub-heading of each column.")

    # Q20 - author contributions
    p.strike(["Q20"], 27, f(27, "LL: .", 2613), "2613",
             'Replace "LL: ." with "LL: Conceptualization, Formal analysis, Methodology, Writing - original draft."')

    # Q23 / Q28 - retracted reference
    p.strike(["Q23", "Q28"], 6, [f(6, "Peng et al.,", 145), f(6, "2024;", 146)], "145-146",
             'Delete "Peng et al., 2024; " (retracted article): "...environmental performance (Xu et al., 2024)."')
    p.strike(["Q23", "Q24", "Q28"], 28, p.block(28, "R", 2741, 2743), "2741-2743",
             "Delete this reference: the article has been retracted (Retraction Note, J. Knowl. Econ., doi: 10.1007/s13132-025-03052-6).")

    # Reference list - Q11 / Q24-Q27 / Q29
    p.strike(["Q11", "Q24", "Q29"], 27, p.block(27, "L", 2599, 2600), "2599-2600",
             "Replace this entry with: Cheng, W., Li, Q., Wu, Q., Ye, F., and Jiang, Y. (2024). Digital capability and green innovation: the perspective of green supply chain collaboration and top management's environmental awareness. Heliyon 10 (11), e32290. doi: 10.1016/j.heliyon.2024.e32290")
    p.strike(["Q11", "Q24", "Q25", "Q29"], 27, p.block(27, "L", 2601, 2602), "2601-2602",
             "Replace this entry with: Xu, Y., and Shi, Q. (2025). Executive financial background and corporate digital transformation: empirical evidence from China. Appl. Econ. Lett. 33 (10), 1653-1661. doi: 10.1080/13504851.2024.2449540")
    p.strike(["Q24", "Q25"], 27, f(27, "(prepublish), 1–36.", 2607), "2607",
             'Replace ",(prepublish), 1-36." with " 31 (3), 2093-2128. doi: 10.1007/s11142-026-09972-1"')
    p.caret(["Q11"], 28, f(28, "DesJardine", 2671), "2671",
            "Insert new reference: Dahiya, M., and Singh, S. (2021). The linkage between CSR and cost of equity: an Indian perspective. Sustainability Account. Manag. Policy J. 12 (3), 499-521.", before=True)
    p.caret(["Q26"], 28, f(28, "Paper.", 2693), "2693",
            "Complete the reference as: Holladay, J. S. (2010). Are Exporters Mother Nature's Best Friends? FREIT Working Paper. Forum for Research in Empirical International Trade (FREIT). Available at: https://cenrep.ncsu.edu/cenrep/wp-content/uploads/2015/08/Holladay.pdf (Accessed 2 October 2026).")
    p.caret(["Q24"], 28, f(28, "30 (5), 580–599.", 2705), "2705", 'Insert " doi: 10.1111/corg.12486"')
    p.strike(["Q26"], 28, [f(28, "NBER Working Paper.", 2715), f(28, "No.26884.", 2716)], "2715-2716",
             'Replace "NBER Working Paper. No.26884." with "NBER Working Paper No. 26884. Cambridge, MA: National Bureau of Economic Research (NBER). doi: 10.3386/w26884"')
    p.caret(["Q24", "Q27"], 28, f(28, "Int. Rev. Econ. and Finance.", 2718), "2718",
            'Insert " 93 (PA), 772-788. doi: 10.1016/j.iref.2024.03.056"')
    p.strike(["Q24", "Q27"], 28, f(28, "groups[EB/OL]. EconPapers.", 2722), "2722",
             'Replace "groups[EB/OL]. EconPapers." with "groups. Statistical Software Components S457877, Boston College Department of Economics. Available at: https://ideas.repec.org/c/boc/bocode/s457877.html" (software module, no DOI/volume/pages).')
    p.caret(["Q24", "Q27"], 28, f(28, "Manag. Sci.", 2734), "2734",
            'Insert " 72 (6), 5171-5192. doi: 10.1287/mnsc.2024.05701"')
    p.caret(["Q11"], 28, f(28, "Orazalin", 2739), "2739",
            "Insert new reference: National Bureau of Statistics of China (2022). China Energy Statistical Yearbook 2022. Beijing: China Statistics Press.", before=True)
    p.caret(["Q11"], 28, f(28, "Sun, B., and Zheng, S. (2024)", 2752), "2752",
            "Insert new reference: Sun, C., and Zeng, Y. (2023). Does the green credit policy affect the carbon emissions of heavily polluting enterprises? Energy Policy 180, 113679. doi: 10.1016/j.enpol.2023.113679", before=True)
    p.caret(["Q29"], 28, f(28, "Yip, W. S. (2025)", 2754), "2754", 'Change "(2025)" to "(2025a)".')
    p.caret(["Q29"], 28, f(28, "Wang, L. (2025)", 2757), "2757", 'Change "(2025)" to "(2025b)".')
    p.strike(["Q11", "Q24", "Q29"], 28, f(28, "Umar, F.,", 2761), "2761-2762",
             'Replace "Umar, F.," with "Farooq, U.," (surname error) and add " doi: 10.1016/j.iref.2023.08.010" at the end of the entry.')
    p.caret(["Q11"], 28, f(28, "Whited, T. M.", 2771), "2771",
            "Insert new reference: Wang, P., and Jin, S. (2025). Can green funds improve corporate carbon performance? Firm-level evidence from China. Sustainability 17 (12), 5409. doi: 10.3390/su17125409", before=True)
    p.caret(["Q29"], 29, f(29, "Zhou, S. (2024)", 2844), "2844", 'Change "(2024)" to "(2024a)".')
    p.caret(["Q29"], 29, f(29, "Zuo, S. (2024)", 2847), "2847", 'Change "(2024)" to "(2024b)".')

    # Additional corrections E1-E7
    p.strike(["E1"], 16, p.block(16, "R", 1346, 1355, start=("Certain external shocks", 1346)), "1346-1355",
             'Delete this duplicated passage ("Certain external shocks ... carbon emissions performance." already appears at l.1337-1346).')
    p.strike(["E2"], 14, f(14, "*, **, ***", 1150), "1150", 'Replace "*, **, ***" with "***, **, *" (*** = 1%, ** = 5%, * = 10%).')
    p.strike(["E2"], 17, f(17, "*, **, ***", 1407), "1407", 'Replace "*, **, ***" with "***, **, *" (*** = 1%, ** = 5%, * = 10%).')
    p.strike(["E3"], 16, [_last(p, 16, "In", 1310), f(16, "Column", 1311)],
             "1310-1311", 'Delete "In Column " so that the sentence reads "Column (2) adds province x year fixed effects...".')
    p.strike(["E4"], 9, _nth(p, 9, "baseline", 541, 1), "541", 'Delete the repeated word "baseline".')
    p.strike(["E5"], 9, f(9, "a individual", 503), "503", 'Replace "a individual" with "an individual".')
    p.strike(["E6"], 10, f(10, "interest.,", 588), "588", 'Replace "interest.," with "interest." (delete the comma).')
    p.caret(["E7"], 16, f(16, "in year t", 1368), "1368", 'Insert a full stop: "...of firm j in year t. NCit refers to..."')
    p.strike(["E7"], 16, p.block(16, "R", 1371, 1372, start=("Cost", 1371), end=("industry i.", 1372)), "1371-1372",
             'The subscripts are transposed. Change to: "Costjt captures the operating costs of firm j in year t, whereas Costit indicates the aggregate operating costs of industry i." (Swap only the subscripts it/jt.)')


def _nth(p, pno, word, line, nth):
    """nth occurrence of a whole word on a line (search_for merges adjacent identical hits)."""
    page = p.page(pno)
    nums = margin_numbers(page)
    hits = [pymupdf.Rect(w[:4]) for w in page.get_text("words")
            if w[4] == word and line_of(page, pymupdf.Rect(w[:4]), nums) == line]
    return hits[nth]


def _last(p, pno, word, line):
    return _nth(p, pno, word, line, -1)


# --------------------------------------------------------------------------------------
# Author queries form
# --------------------------------------------------------------------------------------
CELLS = {  # query -> (page, y0, y1) of its row in the "Author's response" column
    "Q1": (1, 592.1, 617.0), "Q2": (1, 617.3, 711.9), "Q3": (1, 712.2, 767.0),
    "Q4": (2, 102.1, 156.9), "Q5": (2, 157.2, 172.1), "Q6": (2, 172.4, 257.0), "Q7": (2, 257.3, 312.1),
    "Q8": (2, 312.3, 357.2), "Q9": (2, 357.4, 392.3), "Q10": (2, 392.5, 417.4), "Q11": (2, 417.7, 462.5),
    "Q12": (2, 462.8, 497.6), "Q13": (2, 497.9, 512.8), "Q14": (2, 513.1, 557.9), "Q15": (2, 558.2, 652.8),
    "Q16": (2, 653.1, 687.9), "Q17": (2, 688.2, 703.1), "Q18": (2, 703.4, 728.3), "Q19": (2, 728.5, 753.4),
    "Q20": (3, 102.1, 156.9), "Q21": (3, 157.2, 231.9), "Q22": (3, 232.1, 296.9), "Q23": (3, 297.1, 332.0),
    "Q24": (3, 332.3, 397.0), "Q25": (3, 397.3, 422.2), "Q26": (3, 422.4, 447.3), "Q27": (3, 447.6, 472.5),
    "Q28": (3, 472.7, 497.6), "Q29": (3, 497.9, 562.6), "Q30": (3, 562.9, 707.3),
}
X0, X1 = 425.5, 544.0


def fits(text, rect, size):
    scratch = pymupdf.open()
    pg = scratch.new_page(width=rect.width + 20, height=rect.height + 20)
    r = pymupdf.Rect(0, 0, rect.width - 4, rect.height - 3)
    return pg.insert_textbox(r, text, fontsize=size, fontname="helv") >= 0


def freetext(page, rect, text, size):
    # A FreeText annotation renders its /Contents, so the visible text and the content are the same.
    a = page.add_freetext_annot(rect, text, fontsize=size, fontname="helv", text_color=(0, 0, 0.55),
                                fill_color=(1, 1, 0.88), align=0)
    a.set_info(title=AUTHOR)
    a.update()
    return a


def locations(log, q):
    pages = {}
    for queries, pno, lines, *_ in log:
        if q in queries:
            pages.setdefault(pno, [])
            if lines not in pages[pno]:
                pages[pno].append(lines)
    first = lambda ls: int(re.match(r"\d+", ls).group())
    return "; ".join(f"p.{pg} l.{', '.join(sorted(ls, key=first))}" for pg, ls in sorted(pages.items()))


def fill_form(p, responses):
    form_log = []
    for q, (pno, y0, y1) in CELLS.items():
        page = p.page(pno)
        rect = pymupdf.Rect(X0, y0 + 0.8, X1, y1 - 0.8)
        full = responses[q]
        size = next((s for s in (7, 6.5, 6, 5.5) if fits(full, rect, s)), None)
        if size:
            freetext(page, rect, full, size)
            form_log.append((q, "full text in cell", size))
            continue
        # Show the longest run of whole sentences that fits; the full text goes in a note.
        box = pymupdf.Rect(X0, y0, X1 - 14, y1)
        sentences = [x for x in re.split(r"(?<!\d)(?<=\.)[ \t]+|\n+", full) if x.strip()]
        short, size = None, None
        for k in range(len(sentences), 0, -1):
            cand = " ".join(sentences[:k]) + " [continued in the attached comment]"
            size = next((s for s in (7, 6.5, 6, 5.5) if fits(cand, box, s)), None)
            if size:
                short = cand
                break
        if short is None:
            short = "See attached comment."
            size = next(s for s in (7, 6.5, 6, 5.5, 5) if fits(short, box, s))
        freetext(page, pymupdf.Rect(X0, y0 + 0.8, X1 - 14, y1 - 0.8), short, size)
        note = page.add_text_annot(pymupdf.Point(X1 - 12.5, y0 + 1.5), full, icon="Comment")
        note.set_info(title=AUTHOR, content=f"[{q}] Author's response:\n\n{full}", subject=q)
        note.set_colors(stroke=YELLOW)
        note.update()
        form_log.append((q, "summary in cell + full text in comment", size))

    # Alt-text row (no query number) - covered by Q30
    freetext(p.page(3), pymupdf.Rect(X0, 708.5, X1, 740),
             "Alt-Text for Figures 1-4: approved without changes (see response to Q30).", 7)
    page = p.page(4)

    # Additional corrections below the form
    body = "ADDITIONAL AUTHOR CORRECTIONS (not raised in the queries; each is also marked in the proof)\n\n" + \
        "\n".join(f"{e}. {t}" for e, t in ADDITIONAL)
    size = 8
    height = next(h for h in range(60, 400, 4) if fits(body, pymupdf.Rect(50, 0, 545, h), size))
    freetext(page, pymupdf.Rect(50, 280, 545, 280 + height + 6), body, size)
    return form_log


RESPONSE_CHANGES_MD = """\
Apart from the edits below, every response is copied word for word from the Word file. The only general change is that page/line references (p.X, l.Y) were added.

| Query/item | Change | Reason |
|---|---|---|
| Q10 | "4.1.x-4.3.x (tertiary)" -> "4.1.x-4.3.x and 6.2.x-6.5.x (tertiary)" | 6.2.1-6.5.2 in the proof are also third-level headings |
| Q14 | Added the Table 5 caption text "TABLE 5 Instrumental variable (2SLS) estimation results." | The original only said "a proper caption will be added" without giving the caption, so the editor would have to ask again. **Please check that you are happy with this wording** |
| Q14 | Added: on p.17 l.1500, "presented in Table 4" -> "Table 5" | The IV diagnostic statistics are in Table 5; Table 4 is the greenwashing test (a wrong table number) |
| Q29 item 2 | Quote changed to "(Jin et al., 2017; Cheng et al., 2021)" | The original quote did not match the proof sentence, and inserting it as written would have made the sentence ungrammatical |
| Q29 item 7 | Added the existing "Lu, 2025" to the quoted citation cluster | The proof cluster contains Lu, 2025; leaving it out could be read as deleting it |
| Q29 item 9 | "Section 1 and Section 2.2 refer to 2025a" -> "Section 1 refers to 2025a" | Section 2.2 of the proof contains no Tan et al. citation |
| Q29 items 8/10/12-18/21 | Specified the exact insertion point for each citation | The original gave only the section, not where to insert |
"""

E_MAP_MD = """\
| Number in the PDF | Original Word number | Status |
|---|---|---|
| E1 | E1 | Kept (duplicated paragraph) |
| E2 | E2 | Kept; scope changed to the notes of Table 2 and Table 4 (no other table has this note) |
| E3 | E3 | Kept; changed to "Column (2) adds..." (the original "In Column (2) adds" was still ungrammatical) |
| E4 | E4 | Kept |
| E5 | E6 | Kept |
| E6 | E7 | Kept |
| E7 | E8 (i) | Kept (Cost subscripts transposed), plus the missing full stop after "in year t" in the same sentence |
| - | E5 | Removed (duplicates Q11 item 9) |
| - | E8 (ii) | **Removed: it was wrong.** Equation 5 divides by Revenue, so CEjt is an intensity; changing it to "total carbon emissions" would introduce an error |
| - | E9-E12 | Removed (not obvious low-level errors) |
"""


def write_checklist(p, responses, form_log, path):
    rows = []
    for i, (queries, pno, lines, kind, instr) in enumerate(p.log, 1):
        instr = instr.replace("|", "\\|").replace("\n", " ")
        if len(instr) > 220:
            instr = instr[:217] + "..."
        rows.append(f"| {i} | {', '.join(queries)} | {pno} | {lines} | {kind} | {instr} |")
    qrows = []
    for q, mode, size in form_log:
        mode = "Full text in cell" if mode.startswith("full") else "Summary in cell + full text in note"
        qrows.append(f"| {q} | {mode} | {locations(p.log, q) or '- (no in-text change)'} |")
    md = [
        "# Author's proof annotation checklist",
        "",
        "## Files and how to submit",
        "",
        "- `Authors-Proof_annotated.pdf`: the annotated proof to submit. Upload it to the **Production Forum**.",
        "- The annotations are standard PDF comments. In Adobe Acrobat Reader, open the **Comments** panel to see them all.",
        "  - Red strike-out = delete or replace; blue caret = insert; yellow highlight = see comment.",
        "  - Every comment starts with `[Qn]` or `[En]` plus `p.page, l.line`, as the editor requires.",
        "- The proof text is unchanged (the PDF is edit-protected and remains so). The comments are appended to the original file as an incremental save.",
        "- Page numbers are PDF page numbers (1-29); line numbers are the numbers printed in the proof margins.",
        "",
        "## 1. Changes made to the Word responses",
        "",
        RESPONSE_CHANGES_MD,
        "## 2. Additional corrections: numbering and selection",
        "",
        E_MAP_MD,
        "The text of E1-E7 is in the yellow box under the queries form on PDF page 4, and each item is also marked at its location in the article.",
        "",
        "## 3. Query responses (Author's response column on PDF pages 1-4)",
        "",
        "| Query | How it is shown | Locations marked in the article |",
        "|---|---|---|",
        *qrows,
        "",
        "The unnumbered Alt-Text row (PDF page 3) says: approved without changes (see Q30).",
        "",
        f"## 4. In-text annotations ({len(p.log)} in total)",
        "",
        "| # | Query | PDF page | Line(s) | Type | Comment text |",
        "|---|---|---|---|---|---|",
        *rows,
        "",
    ]
    open(path, "w").write("\n".join(md))


def main(proof, response_docx, out_pdf, checklist):
    responses = load_responses(response_docx)
    # Annotate a byte-for-byte copy and append the comments as an incremental update, so the
    # original proof (content, security settings) is left untouched - as Acrobat does.
    shutil.copyfile(proof, out_pdf)
    p = Proof(out_pdf)
    mark_corrections(p, responses)
    form_log = fill_form(p, responses)
    p.doc.save(out_pdf, incremental=True, encryption=pymupdf.PDF_ENCRYPT_KEEP)
    write_checklist(p, responses, form_log, checklist)
    print(f"{len(p.log)} in-text annotations; {len(form_log)} query responses")


if __name__ == "__main__":
    main(*sys.argv[1:5])
