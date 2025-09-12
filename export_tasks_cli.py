#!/usr/bin/env python3
import argparse
import getpass
import logging
import requests


def parse_args():
    parser = argparse.ArgumentParser(description="Export tasks from iQUAVIS")
    parser.add_argument("--base-url", default="http://rdgpm0701", help="Base URL of the iQUAVIS API")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    return parser.parse_args()


def fetch_projects(session: requests.Session, base_url: str) -> list:
    """Fetch project list from the server.

    Args:
        session: requests session with authentication configured.
        base_url: API base URL, e.g., ``http://rdgpm0701``.

    Returns:
        List of projects (parsed JSON).
    """
    projects_url = f"{base_url}/projects"
    logging.debug("Fetching projects from %s", projects_url)
    try:
        response = session.get(projects_url, timeout=30)
        logging.debug("Response status: %s", response.status_code)
        logging.debug("Response headers: %s", dict(response.headers))
        if response.content:
            logging.debug("Response content: %s", response.content[:200])
        response.raise_for_status()
        return response.json()
    except Exception:
        logging.exception("Failed to fetch projects")
        raise


def main():
    args = parse_args()
    logging.basicConfig(level=logging.DEBUG if args.debug else logging.INFO,
                        format="[%(levelname)s] %(message)s")

    print("iQUAVIS login")
    user_id = input("User ID: ")
    password = getpass.getpass("Password: ")

    # For real-world usage, authentication would happen here.
    session = requests.Session()
    session.auth = (user_id, password)
    print("Authenticated successfully.")

    try:
        projects = fetch_projects(session, args.base_url)
        print("Fetched projects:")
        for p in projects:
            print(f"- {p}")
    except Exception as exc:
        print(f"Failed to fetch projects: {exc}")


if __name__ == "__main__":
    main()
