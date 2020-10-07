from ._i_data_gateway import NoExcelFiles
from ._file_data_gateway import OutlookDataGateway as _OutlookDataGateway
from ._dialog_data_gateway import DialogDataGateway as _DialogDataGateway


GATEWAYS = {
    'outlook': _OutlookDataGateway,
    'dialog': _DialogDataGateway
}
