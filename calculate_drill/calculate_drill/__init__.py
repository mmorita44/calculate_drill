from .drill import Drill
import sys

def main():
    return Drill(sys.argv[1]).calculate()
