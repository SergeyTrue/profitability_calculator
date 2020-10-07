from collections import defaultdict
from typing import List, Dict
import importlib

from . import config
from mc.common import User, Role


ALL_ACCEPTABLE_USERS = []


def _parse_users(all_acceptable_users: List[Dict]):
    future_boss = defaultdict(list)  # boss_id: [user1, user2]
    users = {}  # id: User

    for item in all_acceptable_users:
        company = item.get('company', 'NO_COMPANY')
        user = User(name=item['name'], email=item['email'], region=item['region'], role=Role(item['role']),
                    company=company, bosses=[])
        users[item['id']] = user

        bosses_id = item['bosses_id']
        if bosses_id:
            for pk in bosses_id:
                if pk in users:
                    user.bosses.append(users[pk])
                else:
                    future_boss[pk].append(user)

        if item['id'] in future_boss:
            sales_managers = future_boss[item['id']]
            for sale_manager in sales_managers:
                sale_manager.bosses.append(user)

    return list(users.values())


_module = config.ALL_USERS.parent / config.ALL_USERS.stem
_module = _module.parts[-3:]
_module = '.'.join(_module)
_m = importlib.import_module(_module)
ALL_ACCEPTABLE_USERS.extend(_parse_users(all_acceptable_users=getattr(_m, 'USERS')))
