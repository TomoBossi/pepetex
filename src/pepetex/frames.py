import shutil
import tempfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image

import utils
import errors
from extract import extract
from compress import compress
from namespaces import PREFIX_NAMESPACES
