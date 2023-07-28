import jpype
import asposecells

jpype.startJVM()
from asposecells.api import Workbook, ProtectionType

# Load Excel file
workbook = Workbook("workbook.xlsx")

# Protect workbook with desired protection type
workbook.protect(ProtectionType.STRUCTURE, "123456")

# Save protected Excel file
workbook.save("QuanLyKho.xlsm")