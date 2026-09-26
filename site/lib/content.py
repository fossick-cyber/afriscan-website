"""Content files: front matter, Markdown and ::: container components."""
import re
from dataclasses import dataclass, field

import yaml
from markdown_it import MarkdownIt

FRONT = re.compile(r"\A---[ \t]*\n(.*?)\n---[ \t]*\n?(.*)\Z", re.S)
OPEN = re.compile(r"^(?P<fence>:{3,})(?P<name>[a-z][\w-]*)[ \t]*(?:\{(?P<attrs>.*)\})?[ \t]*$")
CLOSE = re.compile(r"^(?P<fence>:{3,})[ \t]*$")
ATTR = re.compile(r'([A-Za-z_][\w-]*)\s*=\s*(?:"([^"]*)"|\'([^\']*)\'|([^\s"\']+))')
HEAD_ID = re.compile(r"<h([2-4])>(.*?)\s*\{#([A-Za-z][\w-]*)\}</h\1>", re.S)
HEAD_PLAIN = re.compile(r"<h([2-4])>(.*?)</h\1>", re.S)
TAG = re.compile(r"<[^>]+>")


class ContentError(Exception):
    pass


def split_front_matter(raw, path):
    m = FRONT.match(raw)
    if not m:
        raise ContentError(f"{path}: missing front matter (--- ... ---) at the top")
    try:
        meta = yaml.safe_load(m.group(1)) or {}
    except yaml.YAMLError as e:
        raise ContentError(f"{path}: front matter is not valid YAML ({e}). Hint: put any value that contains \": \" or starts with a quote, [, {{, *, & or # in double quotes")
    if not isinstance(meta, dict):
        raise ContentError(f"{path}: front matter must be a mapping")
    return meta, m.group(2)


def make_markdown():
    md = MarkdownIt("commonmark", {"html": True, "typographer": True})
    md.enable(["table", "replacements", "smartquotes"])
    md.options["quotes"] = "“”‘’"
    return md


@dataclass
class Block:
    name: str
    attrs: dict
    fence: int
    line: int
    children: list = field(default_factory=list)


def parse_attrs(text, where):
    attrs, pos = {}, 0
    text = (text or "").strip()
    for m in ATTR.finditer(text):
        if text[pos:m.start()].strip():
            raise ContentError(f"{where}: cannot read attributes {text!r}")
        attrs[m.group(1)] = next(g for g in m.groups()[1:] if g is not None)
        pos = m.end()
    if text[pos:].strip():
        raise ContentError(f"{where}: cannot read attributes {text!r}")
    return attrs


def parse_blocks(body, path):
    """Split a body into text runs and nested ::: blocks.

    A bare fence closes the innermost open block and must have the same number of colons as its
    opener. Using more colons for outer blocks is only a readability convention."""
    root = Block("root", {}, 0, 0)
    stack, buf = [root], []

    def flush():
        if buf:
            stack[-1].children.append("\n".join(buf) + "\n")
            buf.clear()

    in_code = False
    for n, line in enumerate(body.split("\n"), 1):
        if line.lstrip().startswith("```"):
            in_code = not in_code
        if not in_code and (m := OPEN.match(line)):
            flush()
            b = Block(m["name"], parse_attrs(m["attrs"], f"{path}:{n}"), len(m["fence"]), n)
            stack[-1].children.append(b)
            stack.append(b)
        elif not in_code and (m := CLOSE.match(line)):
            flush()
            if len(stack) == 1:
                raise ContentError(f"{path}:{n}: closing fence with no open block")
            if len(m["fence"]) != stack[-1].fence:
                raise ContentError(f"{path}:{n}: fence of {len(m['fence'])} colons does not close "
                                   f"':::{stack[-1].name}' (opened with {stack[-1].fence} on line {stack[-1].line})")
            stack.pop()
        else:
            buf.append(line)
    flush()
    if len(stack) > 1:
        raise ContentError(f"{path}: ':::{stack[-1].name}' opened on line {stack[-1].line} is never closed")
    return root.children


def slugify(text):
    text = TAG.sub("", text).lower()
    text = (text.replace("&amp;", "and").replace("ç", "c").replace("ã", "a").replace("á", "a").replace("à", "a")
            .replace("â", "a").replace("é", "e").replace("ê", "e").replace("í", "i").replace("ó", "o")
            .replace("ô", "o").replace("õ", "o").replace("ú", "u"))
    return re.sub(r"[^a-z0-9]+", "-", text).strip("-")[:60] or "section"


def add_heading_ids(html, used):
    def explicit(m):
        hid = m.group(3)
        used.add(hid)
        return f'<h{m.group(1)} id="{hid}">{m.group(2)}</h{m.group(1)}>'

    html = HEAD_ID.sub(explicit, html)

    def auto(m):
        base = slugify(m.group(2))
        hid, i = base, 2
        while hid in used:
            hid, i = f"{base}-{i}", i + 1
        used.add(hid)
        return f'<h{m.group(1)} id="{hid}">{m.group(2)}</h{m.group(1)}>'

    return HEAD_PLAIN.sub(auto, html)


def wrap_tables(html, label="Table"):
    return re.sub(r"<table>(.*?)</table>",
                  lambda m: f'<div class="table-wrap" tabindex="0" role="region" aria-label="{label}">'
                            f'<table>{m.group(1)}</table></div>',
                  html, flags=re.S)
