import xml.etree.ElementTree as ET

def get_svg_html():
    def parse_svg(file):
        tree = ET.parse(file)
        root = tree.getroot()
        paths = []
        for p in root.iter():
            if p.tag.endswith("path"):
                d = p.attrib.get("d", "")
                transform = p.attrib.get("transform", "")
                paths.append((d, transform))
        return paths

    text_paths = parse_svg("text.svg")
    html = "<html><body><svg viewBox=\"0 0 1080 1920\" style=\"width: 500px; height: auto; background: #eee;\">"
    for i, (d, t) in enumerate(text_paths):
        trans = f' transform="{t}"' if t else ''
        html += f"<path d=\"{d}\"{trans} fill=\"black\" id=\"path-{i}\" />\n"
    html += "</svg></body></html>"

    with open("preview_text.html", "w") as f:
        f.write(html)

get_svg_html()
