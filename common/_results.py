from pathlib import Path
from dataclasses import dataclass
from typing import Optional, List, Dict

import pandas as pd

from ._user import User


@dataclass
class ResultLineItem:
    path: Path
    discount_margin_relation_merged_df: Optional[pd.DataFrame]
    discount_margin_relation_grouped_df: Optional[Dict[str, pd.DataFrame]]
    equal_margin_df: Optional[pd.DataFrame]
    price_by_type_of_equipment: Optional[pd.DataFrame]
    bad_cost: Optional[pd.DataFrame]
    full_emulator: Optional[pd.DataFrame]
    complete_price_cost: Optional[pd.DataFrame]
    bad_articles: Optional[pd.DataFrame]
    minimum_quantity_warning: Optional[pd.DataFrame]
    multiplicity_warning: Optional[pd.DataFrame]
    bad_cost_ratio: Optional[float]
    warning_999999_price: Optional[pd.DataFrame]

    def new_to_dsm(self):
        return ResultLineItem(self.path,
                              discount_margin_relation_merged_df=None,
                              discount_margin_relation_grouped_df=None,
                              equal_margin_df=None,
                              price_by_type_of_equipment=self.price_by_type_of_equipment,
                              bad_cost=None,
                              full_emulator=None,
                              complete_price_cost=self.complete_price_cost,
                              bad_articles=self.bad_articles,
                              minimum_quantity_warning=self.minimum_quantity_warning,
                              multiplicity_warning=self.multiplicity_warning,
                              bad_cost_ratio=None,
                              warning_999999_price=self.warning_999999_price,)


class SuccessUserResult:
    def __init__(self, children=None):
        super().__init__()

        if children is None:
            children = []

        self.receiver: Optional[User] = None
        self.sender: Optional[User] = None
        self.children: List[ResultLineItem] = list(children)

    def new_to_dsm(self):
        return SuccessUserResult(children=(item.new_to_dsm() for item in self.children))


class BadResult:
    def __init__(self, wrong):
        self.wrong = wrong


class ConcatenatedUserResult:
    def __init__(self, bad_user_result: BadResult, accepted_user_result: SuccessUserResult):
        self.receiver = None
        self.bad_user_result = bad_user_result
        self.accepted_user_result = accepted_user_result

    #     self._sender = None
    #
    # @property
    # def sender(self):
    #     return self._sender
    #
    # @sender.setter
    # def sender(self, value):
    #     self._sender = value
    #     if self.bad_user_result:
    #         self.bad_user_result.sender = value
    #     if self.accepted_user_result:
    #         self.accepted_user_result.sender = value

    def new_to_dsm(self):
        accepted_user_result = self.accepted_user_result.new_to_dsm() if self.accepted_user_result else None
        bad_user_result = self.bad_user_result

        new = ConcatenatedUserResult(bad_user_result, accepted_user_result)
        return new


class UnknownUserResult:
    def __init__(self, content: str):
        self.content = content


class NoExcelFilesResult:
    def __init__(self, sender: User):
        self.sender = sender
