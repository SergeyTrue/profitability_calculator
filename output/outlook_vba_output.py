from pathlib import Path
from shutil import rmtree
import tempfile

from . import IOutput
from ._result_handler import (SuccessUserResultHandler, SuccessAdminHeartBeatResultHandler, BadResultHandler,
                              UnknownUserResultHandler, ConcatenatedUserResultHandler, NoExcelFilesHandler)
from ..common import Role, SuccessUserResult, ConcatenatedUserResult, UnknownUserResult, NoExcelFilesResult
from .reply_to_mail_thread_adder import (detect_encoding, add_mail_thread_to_final_table, UnknownEncoding,
                                         check_utf_or_win_encoding)


class OutlookVbaOutput(IOutput):

    def __init__(self, dir_to_save: Path, *,
                 boss_filename: str, dsm_filename: str, admin_filename: str, unknown_user_filename: str,
                 initial_mail: Path):
        dir_to_save.mkdir(parents=True, exist_ok=True)
        self.dir_to_save = dir_to_save

        self.admin_attachments_dir = dir_to_save / 'admin_attachments'
        self.boss_attachments_dir = dir_to_save / 'boss_attachments'
        self.dsm_attachments_dir = dir_to_save / 'dsm_attachments'

        for directory in (self.admin_attachments_dir, self.boss_attachments_dir, self.dsm_attachments_dir):
            if directory.is_dir():
                rmtree(str(directory))

        self.admin_attachments_dir.mkdir(exist_ok=True, parents=True)
        self.boss_attachments_dir.mkdir(exist_ok=True, parents=True)
        self.dsm_attachments_dir.mkdir(exist_ok=True, parents=True)

        self.boss_file = dir_to_save / boss_filename
        self.dsm_file = dir_to_save / dsm_filename
        self.admin_file = dir_to_save / admin_filename
        self.unknown_user_file = dir_to_save / unknown_user_filename

        for file in (self.boss_file, self.dsm_file, self.admin_file, self.unknown_user_file):
            file.unlink(missing_ok=True)

        _ = Path(tempfile.gettempdir())
        self.bad_cost_file = _ / 'bad_cost.txt'
        self.boss_email_file = dir_to_save / 'boss_email.txt'
        self.dsm_email_file = dir_to_save / 'dsm_email.txt'
        self.admin_email_file = dir_to_save / 'admin_email.txt'
        self.unknown_email_file = dir_to_save / 'unknown_email.txt'
        self.email_files = {
            Role.BOSS: self.boss_email_file,
            Role.ADMIN: self.admin_email_file,
            Role.DSM: self.dsm_email_file,
            Role.UNKNOWN: self.unknown_email_file

        }



        self._initial_mail = initial_mail

    def process(self, messages):
        for message in messages:
            result = message.data

            file = self.email_files[message.receiver.role]
            with file.open('w', encoding='utf-8') as stream:
                stream.write(message.receiver.email)

            if isinstance(result, NoExcelFilesResult):
                file = self.boss_file if message.receiver.role == Role.BOSS else self.dsm_file
                file.touch()

                handler = NoExcelFilesHandler(file)
                handler.handle(message)

            elif isinstance(result, UnknownUserResult):
                self.unknown_user_file.touch()
                handler = UnknownUserResultHandler(output_file=self.unknown_user_file)
                handler.handle(message)

            elif isinstance(result, SuccessUserResult):
                if message.receiver.role == Role.ADMIN:
                    self._handle_admin(message)

            elif isinstance(result, ConcatenatedUserResult):
                file = self.boss_file if message.receiver.role == Role.BOSS else self.dsm_file
                file.touch()
                dir_ = self.boss_attachments_dir if message.receiver.role == Role.BOSS else self.dsm_attachments_dir

                if result.accepted_user_result:
                    equal_margin_files = tuple((dir_ / f'Сводная {item.path.name}',)
                                               for item in result.accepted_user_result.children)
                else:
                    equal_margin_files = tuple()

                user_success_handler = SuccessUserResultHandler(reply_file=file, equal_margin_xls=equal_margin_files)
                bad_result_handler = BadResultHandler(output_file=file)
                handler = ConcatenatedUserResultHandler(user_success_handler, bad_result_handler)
                #
                # if result.accepted_user_result:
                #     if message.receiver.role == Role.DSM:
                #         result = result.new_to_dsm()
                handler.handle(message)

        # try:
        #     encoding = detect_encoding(self._initial_mail)
        #     print(encoding)
        # except UnknownEncoding:
        #     encoding = check_utf_or_win_encoding(self._initial_mail)
        # if encoding is not None:
        #     for file in (self.boss_file, self.dsm_file, self.admin_file, self.unknown_user_file):
        #         if file.is_file():
        #             with file.open(encoding='utf-8') as stream:
        #                 content = stream.read()
        #                 add_mail_thread_to_final_table(self._initial_mail, encoding, content, file)

    def _handle_admin(self, message):
        result = message.data

        equal_margin_files = tuple((self.admin_attachments_dir / item.path.name,) for item in result.children)
        full_emulator_files = tuple((self.admin_attachments_dir / f'emulator_{item.path.stem}.txt',)
                                    for item in result.children)
        self.admin_file.touch()
        handler = SuccessAdminHeartBeatResultHandler(self.bad_cost_file, self.admin_file,
                                                     full_emulator_files, equal_margin_files)
        handler.handle(message)
