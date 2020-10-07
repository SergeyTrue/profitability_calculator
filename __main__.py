import argparse
from pathlib import Path
from typing import Iterable
import tempfile
import pandas as pd
import datetime
import os
import sys
sys.path.insert(0, r"C:\Users\belose\PycharmProjects\margin_calculation")
from distutils.dir_util import copy_tree
from mc.common import (config, User, UnknownUser, Role, ALL_ACCEPTABLE_USERS, ConcatenatedUserResult,
                       UnknownUserResult, BadResult, NoExcelFilesResult)
from mc.presentation import GATEWAYS, NoExcelFiles
from mc.preparation import DataFramePreparer
from mc.processor import Processor
from mc.router import prepare_messages
from mc.output.outlook_vba_output import OutlookVbaOutput
from mc.output.input_file_modifier_output import InputFileModifierOutput


#print(tempfile.gettempdir())
parser = argparse.ArgumentParser()
subparsers = parser.add_subparsers(dest='data_receipt_way')
subparser1 = subparsers.add_parser('outlook')
subparser1.add_argument('-sender', '--sender_path',
                        default=r'C:\Users\belose\AppData\Local\Temp\margin_calc\sender.txt', type=Path)
subparser1.add_argument('-subj', '--subj_path',
                        default=r'C:\Users\belose\AppData\Local\Temp\margin_calc\subject.txt', type = Path)
subparser1.add_argument('-f_list', '--file_list_path',
                        default=r'C:\Users\belose\AppData\Local\Temp\margin_calc\file_list.txt', type=Path)
subparser1.add_argument('-initial_mail', '--initial_mail_path',
                        default=r'C:\Users\belose\AppData\Local\Temp\margin_calc\InitialMail.html', type=Path)
subparser2 = subparsers.add_parser('dialog')
subparser2.add_argument('-file_list',  type=list)
subparser2.add_argument('-sender', '--sender', default=r'Sergey.Belousov@delaval.com', type=str)
subparser2.add_argument('-subj', '--subj', default=r'test', type=str)

args = parser.parse_args()


def read_sender(users: Iterable[User], sender_path: Path) -> User:
    with sender_path.open() as stream:
        email = stream.read().strip()
        user = next((user for user in users if user.email == email), UnknownUser())
        if isinstance(user, UnknownUser):
            user.email = email
        return user


def read_sender_from_dialog(users: Iterable[User], sender_from_dialog: str) -> User:
    email = sender_from_dialog.strip()
    user = next((user for user in users if user.email == email), UnknownUser())
    if isinstance(user, UnknownUser):
        user.email = email
    return user



def get_admin(users: Iterable[User]) -> User:
    user = next((user for user in users if user.role == Role.ADMIN))
    return user


def read_subject(subject_path: Path) -> str:
    with subject_path.open() as stream:
        value = stream.read().strip()
        return value


args = vars(args)

way = args.pop('data_receipt_way')

gateway = GATEWAYS[way](**args)


input_data = gateway.process(config.valid_suffixes)



if way == 'outlook':

    sender = read_sender(ALL_ACCEPTABLE_USERS, input_data.sender_path)
    print(sender)

    subject = read_subject(input_data.subj_path)
    print(subject)
else:
    sender = read_sender_from_dialog(ALL_ACCEPTABLE_USERS, input_data.sender)
    subject = input_data.subject

admin = get_admin(ALL_ACCEPTABLE_USERS)
print(admin)




prepared = DataFramePreparer().process(input_data, config.sheet_names_to_be_excluded)

messages = []
if isinstance(sender, UnknownUser):
    result = UnknownUserResult(f'Вам нельзя сюда писать')
    unknown_user_messages = prepare_messages(result, sender, admin, subject)
    messages.extend(unknown_user_messages)
else:
    if input_data.file_list:
        bad_user_result = None
        if prepared.wrong:
            bad_user_result = BadResult(wrong=prepared.wrong)

        accepted_user_result = None
        if prepared.accepted:
            accepted_user_result = Processor().process(
                prepared.accepted,
                sender.company,
                minimum_quantity_df=pd.read_csv(str(config.MINIMUM_QUANTITY_FILE), sep='\t'),
                multiplicity_df=pd.read_csv(str(config.MULTIPLICITY_FILE), sep='\t'),
                warning_999999_price=prepared.warning
            )

        result = ConcatenatedUserResult(bad_user_result, accepted_user_result)

        _ = prepare_messages(result, sender, admin, subject)
        messages.extend(_)
    else:
        result = NoExcelFilesResult(sender)
        _ = prepare_messages(result, sender, admin, subject)
        messages.extend(_)

print(messages)

# output = CompositeOutput()
output = OutlookVbaOutput(dir_to_save=Path(tempfile.gettempdir()) / 'margin_calc',
                          boss_filename='boss_file.html',
                          dsm_filename='dsm_file.html',
                          admin_filename='admin_file.html',
                          unknown_user_filename='unknown_user_filename.txt',
                          initial_mail=input_data.initial_mail_path)

modifier = InputFileModifierOutput()
modifier.process(messages)

output.process(messages)


def make_history_path(subject, region, company):

    now = datetime.datetime.now()
    timestamp = str(now.strftime("%Y-%m-%d_%H-%M-%S_"))
    only_num_subj = list([val for val in subject
                    if val.isalpha() or val.isnumeric() or val.isspace()])
    time_and_subject = timestamp + "".join(only_num_subj).strip()

    redion_dir = config.HISTORY_PATH[region]

    full_path = redion_dir + '\\' + 'test' + time_and_subject + company
    try:
        os.makedirs(full_path, exist_ok=True)
        return full_path
    except Exception:
        path_on_my_pc = list(full_path)
        path_on_my_pc[0] = 'C'
        path_on_my_pc = ''.join(path_on_my_pc)
        os.makedirs(path_on_my_pc, exist_ok=True)
        return path_on_my_pc

history_dir = make_history_path(subject, sender.region, sender.company)
print(history_dir)

copy_tree(tempfile.gettempdir() + '\\' + 'margin_calc', history_dir)
