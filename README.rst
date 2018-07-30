Robot Framework Excel Library
=======================================

|Build Status|

Short Description
-----------------

`Robot Framework`_ library for working with Excel documents, based on `openpyxl`_.

Installation
------------

::

    pip install robotframework-excellib

Documentation
-------------

See keyword documentation for robotframework-excellib library: docs_.

Example
-------

.. code:: robotframework

    *** Settings ***
    Library    ExcelLibrary

    *** Test Cases ***
    Check created excel doc
        ${document}=    Create Excel Document    doc_name
        Should Be Equal As Strings    doc_name    ${document}


License
-------

Apache License 2.0

.. _Robot Framework: http://www.robotframework.org

.. _openpyxl: https://pypi.python.org/pypi/openpyxl

.. |Build Status| image:: https://travis-ci.org/peterservice-rnd/robotframework-excellib.svg?branch=master
   :target: https://travis-ci.org/peterservice-rnd/robotframework-excellib

.. _docs: https://rawgit.com/peterservice-rnd/robotframework-excellib/master/docs/ExcelLibrary.html