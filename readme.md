# source and output data are not inclued in this repo

1 a random uuid for flow was created for simapro model within OpenLCA,
which cause the problem that flow referenced by process doesn't correspond
to the right flow with the correct uuid.

2 based on the given mapping file, Sima_pro_background_processes_and_flows
a correct reference between process and flow should be achieved.

# folder simapro
contains all generated xlsx files for simapro models which needs to
be checked for wrong flow uuids.

# sima_pro_background_processes_and_flows.xlsx
contains processes with correct referenced flows

#openpyxl
existing format is not preserved using this library to modify excel file

#xlwings
use this lib to modify excel files if format should be retained 