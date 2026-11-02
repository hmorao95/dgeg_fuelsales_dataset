API reference
=============

Download
--------

.. autofunction:: dgeg_fuelsales.download_latest

.. autofunction:: dgeg_fuelsales.latest_file_url

Updating the workbook
---------------------

.. autofunction:: dgeg_fuelsales.read_dgeg_file

.. autofunction:: dgeg_fuelsales.update_workbook

Building the dataset
--------------------

.. autofunction:: dgeg_fuelsales.read_panel

.. autofunction:: dgeg_fuelsales.build

Result types
------------

.. autoclass:: dgeg_fuelsales.DgegFile
   :members:

.. autoclass:: dgeg_fuelsales.Revision
   :no-members:

.. autoclass:: dgeg_fuelsales.UpdateResult
   :no-members:

.. autoclass:: dgeg_fuelsales.BuildResult
   :no-members:

Constants
---------

.. data:: dgeg_fuelsales.DEFAULT_WORKBOOK

   The source workbook, ``data/processed/dataset_fuel_sales.xlsx``.

.. data:: dgeg_fuelsales.DEFAULT_RAW_DIR

   Where downloaded DGEG files are kept, ``data/raw/dgeg``.

.. data:: dgeg_fuelsales.DEFAULT_OUTPUTS

   Where the dataset is written, ``outputs``.

.. data:: dgeg_fuelsales.MONTHS

   Month names as DGEG writes them, January first.

.. data:: dgeg_fuelsales.OUTPUT_FILES

   The files :func:`~dgeg_fuelsales.build` writes.

Command line
------------

.. autofunction:: dgeg_fuelsales.main
