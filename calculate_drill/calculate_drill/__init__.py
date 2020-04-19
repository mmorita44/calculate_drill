import argparse
from .drill import Drill


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("file", type=str)
    parser.add_argument("first_min", type=int)
    parser.add_argument("first_max", type=int)
    parser.add_argument("second_min", type=int)
    parser.add_argument("second_max", type=int)
    args = parser.parse_args()
    return Drill(args.file).write(args.first_min, args.first_max, args.second_min, args.second_max)
