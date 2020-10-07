from pathlib import Path


RESOURCES = Path(__file__).parents[2] / 'resources'
ALL_USERS = RESOURCES / 'files' / 'users.py'

MINIMUM_QUANTITY_FILE = RESOURCES / r'files\MOQ.txt'
MULTIPLICITY_FILE = RESOURCES / r'files\MULTIPLICITY.txt'

HISTORY_PATH = {
    'central': r'N:\DRM\Центральный',
    'Volgo-Vyatsky': r'N:\DRM\Волго-Вятский',
    'North-West': r'N:\DRM\Северо-Западный',
    'East': r'N:\DRM\Восточный',
    'marketing': r'N:\DRM\Маркетинг',
    'CA': r'N:\DRM\Средняя Азия',
    'misc': r'N:\DRM\Прочее'
}


PRICE_COST_FILE = RESOURCES / r'files\Ultimate_DRC.xlsx'
PRICE_COST_COLUMN = {
    'RU41': ['Customer Price In EUR RU41', 'RU41 DRC'],
    'SE43': ['Customer Price In EUR SE43_KZ10', 'SE43 DRC']
}


sheet_names_to_be_excluded = ('СВОДНАЯ', 'DRM', 'DRM_no_formulas', 'General', 'vgm', 'common', 'расчет', 'расчёт',
                              'general', 'drm', 'drc', 'fgc')
result_sheet_names = ('Общая DRM', 'complete_price_cost_DRM')
valid_suffixes = ['.xlsx']
