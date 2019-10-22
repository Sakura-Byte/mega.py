import random
import os

import pytest

from mega import Mega

TEST_CONTACT = 'test@mega.co.nz'
TEST_PUBLIC_URL = (
    'https://mega.nz/#!hYVmXKqL!r0d0-WRnFwulR_shhuEDwrY1Vo103-am1MyUy8oV6Ps'
)
TEST_FILE = os.path.basename(__file__)
TEST_FOLDER = 'mega.py_testfolder_{0}'.format(random.random())


@pytest.fixture
def mega():
    mega_ = Mega()
    mega_.login(email=os.environ['EMAIL'], password=os.environ['PASS'])
    node = mega_.create_folder(TEST_FOLDER)
    yield mega_
    node_id = node['f'][0]['h']
    mega_.destroy(node_id)


def test_mega(mega):
    assert isinstance(mega, Mega)


def test_login(mega):
    assert isinstance(mega, Mega)


def test_get_user(mega):
    resp = mega.get_user()
    assert isinstance(resp, dict)


def test_get_quota(mega):
    resp = mega.get_quota()
    assert isinstance(int(resp), int)


def test_get_storage_space(mega):
    resp = mega.get_storage_space(mega=True)
    assert isinstance(resp, dict)


def test_get_files(mega):
    files = mega.get_files()
    assert isinstance(files, dict)


def test_get_link(mega):
    file = mega.find(TEST_FILE)
    if file:
        link = mega.get_link(file)
        assert isinstance(link, str)


class TestExport:

    def test_export_folder(self, mega):
        public_url = None
        for _ in range(2):
            result_public_share_url = mega.export(TEST_FOLDER)

            if not public_url:
                public_url = result_public_share_url

            assert result_public_share_url.startswith('https://mega.co.nz/#F!')
            assert result_public_share_url == public_url

    def test_export_single_file(self, mega):
        # Upload a single file into a folder
        folder = mega.find(TEST_FOLDER)
        dest_node_id = folder[1]['h']
        result = mega.upload(
            __file__, dest=dest_node_id, dest_filename='test.py'
        )
        path = f'{TEST_FOLDER}/test.py'
        assert mega.find(path)

        for _ in range(2):
            result_public_share_url = mega.export(path)

            assert result_public_share_url.startswith('https://mega.co.nz/#!')


def test_import_public_url(mega):
    resp = mega.import_public_url(TEST_PUBLIC_URL)
    file_handle = mega.get_id_from_obj(resp)
    resp = mega.destroy(file_handle)
    assert isinstance(resp, int)


def test_create_folder(mega):
    resp = mega.create_folder(TEST_FOLDER)
    assert isinstance(resp, dict)


def test_rename(mega):
    file = mega.find(TEST_FOLDER)
    if file:
        resp = mega.rename(file, TEST_FOLDER)
        assert isinstance(resp, int)


def test_delete_folder(mega):
    folder_node = mega.find(TEST_FOLDER)[0]
    resp = mega.delete(folder_node)
    assert isinstance(resp, int)


def test_delete(mega):
    file = mega.find(TEST_FILE)
    if file:
        resp = mega.delete(file[0])
        assert isinstance(resp, int)


def test_destroy(mega):
    file = mega.find(TEST_FILE)
    if file:
        resp = mega.destroy(file[0])
        assert isinstance(resp, int)


def test_empty_trash(mega):
    # resp None if already empty, else int
    resp = mega.empty_trash()
    if resp is not None:
        assert isinstance(resp, int)


def test_add_contact(mega):
    resp = mega.add_contact(TEST_CONTACT)
    assert isinstance(resp, int)


def test_remove_contact(mega):
    resp = mega.remove_contact(TEST_CONTACT)
    assert isinstance(resp, int)
