from pptx import Presentation
from pptx.util import Inches, Pt


OUTPUT_FILE = "FREQUENT_PATTERN_MINING.pptx"


def set_title(slide, text):
    title = slide.shapes.title
    title.text = text
    for paragraph in title.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(34)
            run.font.bold = True


def add_bullets(slide, lines, left=0.8, top=1.7, width=12.0, height=5.2, font_size=24):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.clear()
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.level = 0
        p.font.size = Pt(font_size)


def add_table(slide):
    headers = ["SID", "Digital-commerce sequence"]
    rows_data = [
        ["S1", "<{Phone}, {Case, ScreenGuard}, {Earbuds}>"],
        ["S2", "<{Laptop}, {Mouse}, {Keyboard}>"],
        ["S3", "<{Phone, Case}, {PowerBank}, {Earbuds}>"],
        ["S4", "<{Laptop}, {Mouse, Keyboard}, {Headphones}>"],
        ["S5", "<{Phone}, {Case}, {Charger}>"],
        ["S6", "<{Laptop}, {Mouse}, {Headphones, Keyboard}>"],
    ]
    table_shape = slide.shapes.add_table(
        len(rows_data) + 1, 2, Inches(0.5), Inches(1.6), Inches(12.2), Inches(4.7)
    )
    table = table_shape.table
    table.columns[0].width = Inches(1.2)
    table.columns[1].width = Inches(11.0)

    table.cell(0, 0).text = headers[0]
    table.cell(0, 1).text = headers[1]

    for r, row in enumerate(rows_data, start=1):
        table.cell(r, 0).text = row[0]
        table.cell(r, 1).text = row[1]

    for r in range(len(rows_data) + 1):
        for c in range(2):
            for p in table.cell(r, c).text_frame.paragraphs:
                p.font.size = Pt(16 if r else 18)
                p.font.bold = r == 0


def build_presentation():
    prs = Presentation()
    layout = prs.slide_layouts[5]

    slides = [
        (
            "FREQUENT PATTERN MINING",
            [
                "University-level presentation",
                "Presented by: Keshav Sharma",
                "Roll No.: 24BCD027",
                "Under the guidance of: Professor Shamim Ahmad",
            ],
        ),
        (
            "Agenda",
            [
                "1. Motivation and definitions",
                "2. Sequence database and support model",
                "3. Core algorithms (GSP, SPADE, PrefixSpan)",
                "4. Support calculations and interpretation",
                "5. Applications, limitations, and conclusions",
            ],
        ),
        (
            "What is Frequent Pattern Mining?",
            [
                "Frequent pattern mining discovers item combinations and sequences that appear repeatedly.",
                "It is a key task in data mining, recommender systems, and customer behavior analytics.",
                "Patterns can be itemsets, subsequences, closed patterns, or maximal patterns.",
                "Goal: find patterns with support >= minimum support threshold.",
            ],
        ),
        (
            "Mathematical Notation",
            [
                "D = {S₁, S₂, ..., Sₙ} is a sequence database with n customer sequences.",
                "S = <e₁, e₂, ..., eₘ> where each event eⱼ is an itemset (eⱼ ⊆ I).",
                "A sequence α is frequent if support(α) >= minsup.",
                "support(α) = |{SID ∈ D : α ⊑ SID}| / |D|",
            ],
        ),
        (
            "Problem Statement and Support Principle",
            [
                "Given D, discover all frequent sequences without enumerating every possible subsequence.",
                "Sequence-level support counts distinct SIDs only (not repeated matches in one SID).",
                "Anti-monotonic property: if α is infrequent, every super-sequence of α is infrequent.",
                "This property enables pruning and scalable mining.",
            ],
        ),
        ("Six-Customer Digital-Commerce Sequence Database", []),
        (
            "Support Snapshot (1-sequences)",
            [
                "|D| = 6 SIDs, use minsup count = 2 (33.33%).",
                "support(<{Phone}>) = 3/6 = 50%  [S1, S3, S5]",
                "support(<{Laptop}>) = 3/6 = 50% [S2, S4, S6]",
                "support(<{Case}>) = 3/6 = 50%   [S1, S3, S5]",
                "All above are frequent because support count >= 2 SIDs.",
            ],
        ),
        (
            "Candidate Generation and Pruning",
            [
                "Join step: combine frequent (k−1)-sequences sharing a common prefix.",
                "Prune step: remove candidate if any (k−1)-subsequence is infrequent.",
                "This is the Apriori-style strategy for sequential pattern mining.",
                "It reduces search space before expensive database scans.",
            ],
        ),
        (
            "GSP (Generalized Sequential Pattern) Algorithm",
            [
                "Level-wise mining: L1 -> L2 -> ... -> Lk.",
                "Each level scans D to count support of generated candidates Ck.",
                "Produces complete frequent sequence set but can be scan-heavy.",
                "Useful baseline for understanding sequence mining mechanics.",
            ],
        ),
        (
            "SPADE Algorithm (Vertical Format)",
            [
                "Transforms D into id-lists: item -> (SID, event-id) occurrences.",
                "Support counting via temporal joins/intersections of id-lists.",
                "Efficient with equivalence classes and reduced candidate handling.",
                "Often faster than repeated horizontal database scans.",
            ],
        ),
        (
            "PrefixSpan (Pattern Growth)",
            [
                "Avoids explicit candidate generation by prefix-based projection.",
                "For each frequent prefix, build projected database and grow recursively.",
                "Works well on long patterns and sparse transactional behavior.",
                "Common practical choice for scalable sequence mining.",
            ],
        ),
        (
            "Support Calculation 1: <{Laptop}, {Mouse}>",
            [
                "Pattern appears in S2: <{Laptop}, {Mouse}, {Keyboard}>.",
                "Pattern appears in S4: <{Laptop}, {Mouse, Keyboard}, {Headphones}>.",
                "Pattern appears in S6: <{Laptop}, {Mouse}, {Headphones, Keyboard}>.",
                "support = 3/6 = 50% (SID count = 3), so pattern is frequent.",
            ],
        ),
        (
            "Support Calculation 2: <{Phone}, {Case}>",
            [
                "S1 satisfies order: {Phone} followed by {Case, ScreenGuard}.",
                "S5 satisfies order: {Phone} followed by {Case}.",
                "S3 has {Phone, Case} together in one event, not as two ordered events.",
                "support = 2/6 = 33.33% (SID count = 2), still frequent at minsup count 2.",
            ],
        ),
        (
            "Support Calculation 3: <{Laptop}, {Mouse}, {Keyboard}>",
            [
                "S2 satisfies: {Laptop} -> {Mouse} -> {Keyboard}.",
                "S6 satisfies: {Laptop} -> {Mouse} -> {Headphones, Keyboard}.",
                "S4 does not satisfy three-event order because {Mouse, Keyboard} are in one event.",
                "support = 2/6 = 33.33% (SID count = 2), frequent at minsup count 2.",
            ],
        ),
        (
            "Closed and Maximal Sequential Patterns",
            [
                "Closed pattern: no super-sequence has identical support.",
                "Maximal pattern: no frequent super-sequence exists.",
                "These condensed representations reduce redundancy in output.",
                "Example: <{Laptop}, {Mouse}, {Keyboard}> is more informative than shorter prefixes alone.",
            ],
        ),
        (
            "Constraint-Based Sequential Mining",
            [
                "Length constraints: minimum/maximum sequence length.",
                "Gap constraints: control allowed distance between matched events.",
                "Item constraints: force inclusion/exclusion of specific products.",
                "Business constraints improve actionability and runtime performance.",
            ],
        ),
        (
            "Applications in Digital Commerce",
            [
                "Cross-sell bundles: discover strong follow-up product sequences.",
                "Personalization: recommend next-best items from frequent paths.",
                "Inventory planning: align stock with multi-step purchase flows.",
                "Campaign timing: trigger offers at high-probability transition points.",
            ],
        ),
        (
            "Limitations and Conclusion",
            [
                "Challenges: large candidate spaces, long sequences, and evolving user behavior.",
                "Support alone may miss rare but high-value patterns; combine with utility metrics.",
                "Frequent pattern mining remains foundational for sequence intelligence.",
                "Key takeaway: always interpret sequence support as distinct SID count over D.",
            ],
        ),
    ]

    for index, (title, bullets) in enumerate(slides, start=1):
        slide = prs.slides.add_slide(layout)
        set_title(slide, title)
        if index == 6:
            add_table(slide)
        else:
            add_bullets(slide, bullets)

    assert len(prs.slides) == 18, f"Expected 18 slides, got {len(prs.slides)}"
    prs.save(OUTPUT_FILE)


if __name__ == "__main__":
    build_presentation()
