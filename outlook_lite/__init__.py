"""A local mail corpus from classic Outlook, for a coding agent to search.

`render` is pure Python and runs anywhere. `dump` and `draft` talk to
Outlook through COM and run only on Windows with classic Outlook open;
both take the Outlook application object as an argument so the tests
can pass a fake.
"""

__version__ = "0.1.0"
