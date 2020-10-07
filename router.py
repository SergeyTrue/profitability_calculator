from typing import Union, Any
from dataclasses import dataclass

from .common import User, Role, UnknownUser, ConcatenatedUserResult, UnknownUserResult, NoExcelFilesResult


@dataclass
class Message:
    sender: User
    receiver: User
    data: Any
    subject: str


def prepare_messages(result, sender: Union[User, UnknownUser], admin: User, subject: str):
    messages = []

    if sender.role == Role.ADMIN:
        sender.role = Role.BOSS

    if isinstance(result, NoExcelFilesResult):
        message = Message(sender=sender, receiver=sender, data=result, subject=subject)
        messages.append(message)
        return messages

    if isinstance(result, UnknownUserResult):
        message = Message(sender=sender, receiver=sender, data=result, subject=subject)
        messages.append(message)
        return messages

    elif isinstance(result, ConcatenatedUserResult):
        if sender.role == Role.BOSS:
            message = Message(sender=sender, receiver=sender, data=result, subject=subject)
            messages.append(message)
        else:
            complete_messages = []
            if result.accepted_user_result is not None:
                for boss in sender.bosses:
                    message = Message(sender=sender, receiver=boss, data=result, subject=subject)
                    complete_messages.append(message)

                messages.extend(complete_messages)

            truncated_result = result.new_to_dsm()
            truncated_message = Message(sender=sender, receiver=sender, data=truncated_result, subject=subject)
            messages.append(truncated_message)

        for_admin_message = Message(sender=sender, receiver=admin, data=result.accepted_user_result, subject=subject)
        messages.append(for_admin_message)
        return messages
