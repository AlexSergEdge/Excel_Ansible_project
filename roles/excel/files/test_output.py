#! /usr/bin/env python
import sys
import json
import pytest

json_data = sys.argv[1]
assert 'Successfully called simple module' in json_data