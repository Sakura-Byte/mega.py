**NOTICE**: If you're reading this on GitHub.com please be aware this is
a mirror of the primary remote located at https://code.richard.do/richardARPANET/mega.py_. Please direct issues and
pull requests there.

--------------

.. _megapy:

Mega.py
=======

|Build Status| |Downloads|

Python library for the Mega.co.nz API, currently supporting:

-  login
-  uploading
-  downloading
-  deleting
-  searching
-  sharing
-  renaming
-  moving files

This is a work in progress, further functionality coming shortly.

For more detailed information see API_INFO.md

How To Use
----------

.. _install-megapy-package:

Install mega.py package
~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   #Run the following command, or run setup from the latest github source
   pip install mega.py

.. _import-megapy:

Import mega.py
~~~~~~~~~~~~~~

.. code:: python

   from mega import Mega

.. _create-an-instance-of-megapy:

Create an instance of Mega.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   mega = Mega()

Login to Mega
~~~~~~~~~~~~~

.. code:: python

   m = mega.login(email, password)
   # login using a temporary anonymous account
   m = mega.login()

Get user details
~~~~~~~~~~~~~~~~

.. code:: python

   details = m.get_user()

Get account balance (Pro accounts only)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   balance = m.get_balance()

Get account disk quota
~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   quota = m.get_quota()

Get account storage space
~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   # specify unit output kilo, mega, gig, else bytes will output
   space = m.get_storage_space(kilo=True)

Get account files
~~~~~~~~~~~~~~~~~

.. code:: python

   files = m.get_files()

Upload a file, and get its public link
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   file = m.upload('myfile.doc')
   m.get_upload_link(file)
   # see mega.py for destination and filename options

Export a file or folder
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   public_exported_web_link = m.export('myfile.doc')
   public_exported_web_link = m.export('my_mega_folder/my_sub_folder_to_share')
   # e.g. https://mega.nz/#F!WlVl1CbZ!M3wmhwZDENMNUJoBsdzFng

Find a file or folder
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   folder = m.find('my_mega_folder')
   # Excludes results which are in the Trash folder (i.e. deleted)
   folder = m.find('my_mega_folder', exclude_deleted=True)

Upload a file to a destination folder
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   folder = m.find('my_mega_folder')
   m.upload('myfile.doc', folder[0])

Download a file from URL or file obj, optionally specify destination folder
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   file = m.find('myfile.doc')
   m.download(file)
   m.download_url('https://mega.co.nz/#!utYjgSTQ!OM4U3V5v_W4N5edSo0wolg1D5H0fwSrLD3oLnLuS9pc')
   m.download(file, '/home/john-smith/Desktop')
   # specify optional download filename (download_url() supports this also)
   m.download(file, '/home/john-smith/Desktop', 'myfile.zip')

Import a file from URL, optionally specify destination folder
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   m.import_public_url('https://mega.co.nz/#!utYjgSTQ!OM4U3V5v_W4N5edSo0wolg1D5H0fwSrLD3oLnLuS9pc')
   folder_node = m.find('Documents')[1]
   m.import_public_url('https://mega.co.nz/#!utYjgSTQ!OM4U3V5v_W4N5edSo0wolg1D5H0fwSrLD3oLnLuS9pc', dest_node=folder_node)

Create a folder
~~~~~~~~~~~~~~~

.. code:: python

   m.create_folder('new_folder')
   m.create_folder('new_folder/sub_folder/subsub_folder')

Returns a dict of folder node name and node_id, e.g.

.. code:: python

    {
      'new_folder': 'qpFhAYwA',
      'sub_folder': '2pdlmY4Z',
      'subsub_folder': 'GgMFCKLZ'
    }

Rename a file or a folder
~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: python

   file = m.find('myfile.doc')
   m.rename(file, 'my_file.doc')


~

.. _`https://code.richard.do/richardARPANET/mega.py`: https://code.richard.do/richardARPANET/mega.py
.. _`https://github.com/meganz/sdk`: https://github.com/meganz/sdk

.. |Build Status| image:: https://travis-ci.org/richardARPANET/mega.py.png?branch=master
   :target: https://travis-ci.org/richardARPANET/mega.py
.. |Downloads| image:: https://pypip.in/d/mega.py/badge.png
   :target: https://crate.io/packages/mega.py/
