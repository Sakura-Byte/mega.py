.. :changelog:

Release History
===============

1.0.2 (2019-11-07)
------------------

- Reverts, "Replace pycrypto dependency with pycryptodome" as breaks login process.


1.0.1 (2019-11-06)
------------------

- When a request fails due to EAGAIN response, retry with exp backoff up to 20 seconds.
- Adds logging, removes print statements.
- Replace pycrypto dependency with pycryptodome.
- Removes Python 2 specific code.


1.0.0 (2019-10-31)
------------------

- Removes broken method ``get_contacts()``.
- Adds support for login with a v2 Mega user account.
- Adds ``export()`` method to share a file or folder, returning public share URL with key.
- Adds code, message attrs to RequestError exception, makes message in raised exceptions include more details.
- Alters ``create_folder()`` to accept a path including multiple sub directories, adds support to create them all (similar to 'mkdir -p' on unix systems).
- Adds ``exclude_deleted=True`` optional arg to ``find()`` method, to exclude deleted nodes from results.

0.9.20 (2019-10-17)
-------------------

- Python 3 bugfix to ``upload`` method.

0.9.19 (2019-10-16)
-------------------

- Python 3 support and bugfixes.
- Update packaging code.
- Added changelog.

0.9.18 (2013-07-04)
-------------------

- Unknown
