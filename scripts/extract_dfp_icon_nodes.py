# -*- coding: utf-8 -*-
"""One-off: extract Lucide icon nodes from nobel-project-hub on Ubuntu VM."""
from __future__ import print_function

import json
import os
import re
import subprocess
import sys

REMOTE = "bimbotubuntu-virtual-machine"
REMOTE_BASE = (
    "/home/bimbot-ubuntu/apps/nobel-project-hub/node_modules/"
    "lucide-react/dist/esm/icons"
)

FUNCTION_ICONS = {
    "1.01": "door-open",
    "1.02": "arrow-left-right",
    "1.03": "refresh-cw",
    "1.04": "circle-dot",
    "1.05": "flame",
    "1.07": "brick-wall",
    "1.08": "shield",
    "1.09": "circle-slash",
    "1.10": "wind",
    "1.11": "air-vent",
    "2.1": "zap",
    "2.2": "waves",
    "2.3": "siren",
    "2.4": "rotate-cw",
    "2.6": "wrench",
    "2.7": "magnet",
    "2.8": "pin",
    "2.9": "door-closed",
    "3.1": "lock",
    "3.2": "key-round",
    "3.3": "key",
    "4.1": "credit-card",
    "4.2": "arrow-right-left",
    "4.3": "scan-line",
    "4.7": "cpu",
    "5.1": "circle-dot",
    "5.2": "nfc",
    "5.3": "fingerprint-pattern",
    "5.4": "grip",
    "6.1": "clock",
    "6.2": "sun",
    "6.3": "combine",
    "6.4": "link-2",
    "6.5": "calendar",
    "7.1": "bell",
    "7.2": "phone",
    "7.3": "layout-grid",
    "7.4": "gauge",
    "fallback": "bell-ring",
}


def _js_node_to_json(text):
    m = re.search(r"const __iconNode = (\[[\s\S]*?\]);", text)
    if not m:
        raise ValueError("icon node not found")
    node = m.group(1)
    node = re.sub(r"(\w+):", r'"\1":', node)
    node = node.replace("'", '"')
    return json.loads(node)


def main():
    out = {}
    for code, name in FUNCTION_ICONS.items():
        remote_path = "{}/{}.js".format(REMOTE_BASE, name)
        proc = subprocess.Popen(
            ["ssh", REMOTE, "cat", remote_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        stdout, stderr = proc.communicate()
        if proc.returncode != 0:
            print("Failed {}: {}".format(name, stderr), file=sys.stderr)
            sys.exit(1)
        text = stdout.decode("utf-8")
        out[code] = {"name": name, "node": _js_node_to_json(text)}

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    dest = os.path.join(root, "lib", "cde", "assets", "dfp_icon_nodes.json")
    assets_dir = os.path.dirname(dest)
    if not os.path.isdir(assets_dir):
        os.makedirs(assets_dir)
    with open(dest, "w", encoding="utf-8") as fh:
        json.dump(out, fh, indent=2)
    print("Wrote {}".format(dest))


if __name__ == "__main__":
    main()
