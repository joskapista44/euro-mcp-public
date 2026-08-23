"""Replacing text where OOXML actually stores it: across run boundaries.

Word and PowerPoint have the same trap in two dialects. A sentence a reader sees
as one line is stored as SEVERAL runs - split for formatting, spell-check state,
revision ids - so "12 500 000" can be `<w:t>12 5</w:t><w:t>00 000</w:t>` (docx)
or `<a:t>12 5</a:t><a:t>00 000</a:t>` (pptx) with nothing visibly wrong. A naive
string replace over the XML then finds nothing and reports success, which is the
worst outcome available: the caller believes the amount was updated.

So the unit of work is the PARAGRAPH. Its runs are concatenated, the search
happens on that text, and the replacement is written back into the runs the match
actually covered: the first one receives the new text, the rest lose the covered
characters. The first run's formatting therefore wins for the whole replacement -
a deliberate choice, because the alternative (splitting runs to preserve
mid-match formatting changes) produces documents that are hard to predict and
harder to review.

The only difference between the two dialects is which element holds the text, so
that is the one thing the caller passes in. Keeping ONE copy of this algorithm is
the point: two copies of the subtlest code in the toolkit would drift, and the
drift would show up as a silently-missed replacement.

Everything outside the touched text elements is left byte-for-byte alone.
"""

# XML entities that appear inside a text element. Text is compared and written
# unescaped, so a search for "A & B" matches what a reader sees, not "A &amp; B".
ENTITIES = (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'), ("&apos;", "'"))


def unescape(text):
    for entity, char in ENTITIES:
        text = text.replace(entity, char)
    return text


def escape(text):
    text = text.replace("&", "&amp;")
    for entity, char in ENTITIES[1:]:
        text = text.replace(char, entity)
    return text


def runs_text(xml, text_re):
    """What a reader sees in this fragment, with the run boundaries collapsed."""
    return "".join(unescape(m.group(2)) for m in text_re.finditer(xml))


def replace_in_runs(xml, find, replace, text_re):
    """Replace every occurrence of `find` in one paragraph. Returns (xml, count).

    Works on the paragraph's concatenated text and writes back per run, so a match
    split across runs is found and replaced exactly once - the case a plain string
    replace over the XML silently misses. An empty `find` is the caller's error and
    is left to the caller to reject: this module does not know which format's error
    type to raise.
    """
    runs = list(text_re.finditer(xml))
    if not runs:
        return xml, 0

    texts = [unescape(m.group(2)) for m in runs]
    joined = "".join(texts)
    if find not in joined:
        return xml, 0

    # Every occurrence, as ranges over the concatenated text.
    spans, at = [], joined.find(find)
    while at != -1:
        spans.append((at, at + len(find)))
        at = joined.find(find, at + len(find))

    # Each run is rebuilt from its OWN slice of the original text: parts no match
    # covered stay as they were, covered parts are dropped, and the run where a
    # match STARTS receives the replacement. Rebuilding (rather than editing in
    # place) is what keeps several matches in one paragraph, and a match spanning
    # two runs, from interfering with each other.
    new_texts, offset = [], 0
    for text in texts:
        run_start, run_end = offset, offset + len(text)
        offset = run_end
        pieces, cursor = [], run_start
        for span_start, span_end in spans:
            if span_end <= run_start or span_start >= run_end:
                continue
            keep_until = min(span_start, run_end)
            if cursor < keep_until:
                pieces.append(joined[cursor:keep_until])
            if run_start <= span_start < run_end:
                pieces.append(replace)
            cursor = max(cursor, min(span_end, run_end))
        if cursor < run_end:
            pieces.append(joined[cursor:run_end])
        new_texts.append("".join(pieces))

    out, index, last = [], 0, 0
    for match in runs:
        out.append(xml[last:match.start()])
        out.append(match.group(1) + escape(new_texts[index]) + match.group(3))
        last = match.end()
        index += 1
    out.append(xml[last:])
    return "".join(out), len(spans)


def write_first_run(xml, text, text_re):
    """Put `text` into the FIRST run of a fragment and empty the rest.

    Used for cell writes, where the target is addressed rather than searched.
    Returns None when there is no run to write into - what to do about that is
    format-specific, so the caller decides.
    """
    runs = list(text_re.finditer(xml))
    if not runs:
        return None
    out, last, index = [], 0, 0
    for match in runs:
        out.append(xml[last:match.start()])
        out.append(match.group(1) + (escape(text) if index == 0 else "") + match.group(3))
        last = match.end()
        index += 1
    out.append(xml[last:])
    return "".join(out)
