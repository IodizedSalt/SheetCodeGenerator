import unittest
import Run

class testCode(unittest.TestCase):
    def test(self):
        self.assertTrue("JV310AIFAPHLU", Run.concatination())
        self.assertTrue("JV310AIFAPHLN", Run.concatination())

def main():
    unittest.main()

if __name__ == "__main__":
    main()