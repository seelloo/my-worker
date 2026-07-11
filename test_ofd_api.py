import sys
import easyofd
print(dir(easyofd))
try:
    with open("test.ofd", "wb") as f:
        pass
    import easyofd.ofd
    print("easyofd.ofd attributes:", dir(easyofd.ofd))
except Exception as e:
    print(e)
