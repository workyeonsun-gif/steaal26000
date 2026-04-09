"""
Chrome Automation Script
- Google Drive files and Notion pages to open automatically in Chrome tabs.
- Supports up to 5 URLs total (Google Drive + Notion combined).

Usage:
    python chrome_opener.py              # Open all configured URLs
    python chrome_opener.py --list       # Show configured URLs
    python chrome_opener.py --add URL    # Add a URL
    python chrome_opener.py --remove N   # Remove URL at index N
    python chrome_opener.py --clear      # Remove all URLs
"""

import json
import os
import platform
import subprocess
import sys
import webbrowser
import time

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
MAX_URLS = 5


def load_config():
    """Load URL list from config file."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"urls": []}


def save_config(config):
    """Save URL list to config file."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def get_chrome_path():
    """Detect Chrome executable path based on OS."""
    system = platform.system()

    if system == "Windows":
        candidates = [
            os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
        ]
    elif system == "Darwin":  # macOS
        candidates = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ]
    else:  # Linux
        candidates = [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium-browser",
            "/usr/bin/chromium",
        ]

    for path in candidates:
        if os.path.exists(path):
            return path

    return None


def open_urls_in_chrome(urls):
    """Open all URLs in Chrome, each in a separate tab."""
    if not urls:
        print("No URLs configured. Use --add to add URLs first.")
        return

    chrome_path = get_chrome_path()

    if chrome_path:
        # First URL opens Chrome (or a new window), rest open as tabs
        cmd = [chrome_path] + urls
        subprocess.Popen(cmd)
        print(f"Chrome opened with {len(urls)} tab(s).")
    else:
        # Fallback: use webbrowser module
        print("Chrome not found at default path. Using system default browser.")
        for i, url in enumerate(urls):
            webbrowser.open(url, new=0 if i == 0 else 2)
            time.sleep(0.3)
        print(f"Opened {len(urls)} tab(s) in default browser.")


def classify_url(url):
    """Classify a URL as Google Drive, Notion, or Other."""
    if "drive.google.com" in url or "docs.google.com" in url:
        return "Google Drive"
    elif "notion.so" in url or "notion.site" in url:
        return "Notion"
    return "Other"


def list_urls(config):
    """Display all configured URLs."""
    urls = config.get("urls", [])
    if not urls:
        print("No URLs configured.")
        return

    print(f"\nConfigured URLs ({len(urls)}/{MAX_URLS}):")
    print("-" * 60)
    for i, url in enumerate(urls):
        tag = classify_url(url)
        print(f"  [{i}] [{tag}] {url}")
    print()


def add_url(config, url):
    """Add a URL to the config."""
    urls = config.get("urls", [])

    if len(urls) >= MAX_URLS:
        print(f"Cannot add more URLs. Maximum is {MAX_URLS}.")
        return config

    tag = classify_url(url)
    if tag == "Other":
        print(f"Warning: URL is not a Google Drive or Notion link.")
        print(f"  Only Google Drive and Notion URLs are intended.")
        answer = input("  Add anyway? (y/n): ").strip().lower()
        if answer != "y":
            print("  Cancelled.")
            return config

    urls.append(url)
    config["urls"] = urls
    save_config(config)
    print(f"Added [{tag}]: {url}")
    print(f"Total: {len(urls)}/{MAX_URLS}")
    return config


def remove_url(config, index):
    """Remove a URL by index."""
    urls = config.get("urls", [])

    if index < 0 or index >= len(urls):
        print(f"Invalid index. Valid range: 0 ~ {len(urls) - 1}")
        return config

    removed = urls.pop(index)
    config["urls"] = urls
    save_config(config)
    print(f"Removed: {removed}")
    print(f"Remaining: {len(urls)}/{MAX_URLS}")
    return config


def clear_urls(config):
    """Remove all URLs."""
    config["urls"] = []
    save_config(config)
    print("All URLs cleared.")
    return config


def interactive_setup():
    """Interactive first-time setup."""
    print("\n=== Chrome Automation - Initial Setup ===")
    print(f"Enter URLs to open automatically (max {MAX_URLS}).")
    print("Supported: Google Drive files, Notion pages")
    print("Type 'done' when finished.\n")

    config = {"urls": []}
    count = 0

    while count < MAX_URLS:
        url = input(f"  URL [{count + 1}/{MAX_URLS}] (or 'done'): ").strip()
        if url.lower() == "done":
            break
        if not url.startswith("http"):
            print("    Invalid URL. Must start with http:// or https://")
            continue
        config["urls"].append(url)
        tag = classify_url(url)
        print(f"    Added [{tag}]")
        count += 1

    save_config(config)
    print(f"\nSetup complete. {count} URL(s) saved to config.json")
    return config


def main():
    args = sys.argv[1:]

    # Handle CLI arguments
    if "--list" in args:
        config = load_config()
        list_urls(config)
        return

    if "--add" in args:
        idx = args.index("--add")
        if idx + 1 >= len(args):
            print("Usage: python chrome_opener.py --add <URL>")
            return
        config = load_config()
        add_url(config, args[idx + 1])
        return

    if "--remove" in args:
        idx = args.index("--remove")
        if idx + 1 >= len(args):
            print("Usage: python chrome_opener.py --remove <index>")
            return
        config = load_config()
        remove_url(config, int(args[idx + 1]))
        return

    if "--clear" in args:
        config = load_config()
        clear_urls(config)
        return

    if "--help" in args or "-h" in args:
        print(__doc__)
        return

    # Default: open URLs
    config = load_config()
    urls = config.get("urls", [])

    # First run: interactive setup
    if not urls:
        config = interactive_setup()
        urls = config.get("urls", [])
        if not urls:
            print("No URLs configured. Exiting.")
            return
        print()

    print("Opening in Chrome...")
    list_urls(config)
    open_urls_in_chrome(urls)


if __name__ == "__main__":
    main()
