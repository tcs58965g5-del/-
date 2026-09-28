#!/usr/bin/env python3
import sys
import ddddocr

def solve(image_path: str) -> str:
    ocr = ddddocr.DdddOcr(show_ad=False)
    with open(image_path, 'rb') as f:
        image_bytes = f.read()
    result = ocr.classification(image_bytes)
    return result.strip()

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('ERROR: no image path provided', file=sys.stderr)
        sys.exit(1)
    try:
        print(solve(sys.argv[1]))
    except Exception as e:
        print(f'ERROR: {e}', file=sys.stderr)
        sys.exit(1)
