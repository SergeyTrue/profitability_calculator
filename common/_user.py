from enum import Enum
from typing import List
from dataclasses import dataclass


class Role(Enum):
    BOSS = 'BOSS'
    DSM = 'DSM'
    ADMIN = 'ADMIN'
    UNKNOWN = 'UNKNOWN'


@dataclass
class User:
    name: str
    email: str
    region: str
    role: Role
    company: str
    bosses: List['User']


class UnknownUser:
    role = Role.UNKNOWN

    def __init__(self):
        self.email = None

