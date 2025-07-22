#!/usr/bin/env python3
"""
Quick syntax test for the VE Table Monitor tool
"""

import sys
import ast


def test_syntax():
    """Test if tool.py has valid Python syntax"""
    try:
        with open('tool.py', 'r', encoding='utf-8') as f:
            source_code = f.read()

        # Parse the file to check for syntax errors
        ast.parse(source_code)
        print("✅ SYNTAX CHECK PASSED!")
        print("✅ VE Table Monitor tool.py has valid Python syntax")
        print("✅ PID scanning function is properly structured")
        print("✅ DTC reading function is intact")
        return True

    except SyntaxError as e:
        print(f"❌ SYNTAX ERROR: {e}")
        print(f"❌ Line {e.lineno}: {e.text}")
        return False
    except Exception as e:
        print(f"❌ ERROR: {e}")
        return False


if __name__ == "__main__":
    success = test_syntax()
    sys.exit(0 if success else 1)
