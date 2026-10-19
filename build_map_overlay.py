"""Create a local raster pyramid from the georeferenced Igarka map.

TIFF strips are decoded with libtiff so the 810-Mpixel source never has to be
loaded into RAM in its entirety. The source GeoTIFF remains unchanged.
"""
import ctypes
import ctypes.util
import json
import math
from pathlib import Path

import numpy as np
import tifffile
from PIL import Image

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / 'Карты' / 'Igarka_Maps.tif'
OUT = ROOT / 'map_overlay'
TILE = 1024
VERSION = 1


def build_map_overlay():
    if not SOURCE.exists():
        return None
    stamp = {'bytes': SOURCE.stat().st_size, 'mtime_ns': SOURCE.stat().st_mtime_ns,
             'builder_version': VERSION}
    metadata_path = OUT / 'metadata.json'
    if metadata_path.exists():
        metadata = json.loads(metadata_path.read_text())
        if metadata.get('source_stamp') == stamp:
            return metadata
    OUT.mkdir(parents=True, exist_ok=True)
    with tifffile.TiffFile(str(SOURCE)) as source:
        page = source.pages[0]
        width, height = page.imagewidth, page.imagelength
        assert page.shape == (height, width, 3) and page.dtype == np.uint8
        assert not page.is_tiled and int(page.planarconfig) == 1
        assert int(page.tags.get(274).value if 274 in page.tags else 1) == 1
        keys = page.tags[34735].value
        crs = next(keys[i+3] for i in range(4, len(keys), 4) if keys[i] == 3072)
        assert crs == 3857, 'Reprojection required: source is not Web Mercator'
        dx, dy = page.tags[33550].value[:2]
        tie = page.tags[33922].value
        west, north = tie[3] - tie[0] * dx, tie[4] + tie[1] * dy
        east, south = west + width * dx, north - height * dy
        rows = page.rowsperstrip
        strip_count = len(page.dataoffsets)

    library_path = ctypes.util.find_library('tiff')
    if not library_path:
        library_path = '/usr/local/opt/libtiff/lib/libtiff.dylib'
    lib = ctypes.CDLL(library_path)
    lib.TIFFOpen.argtypes = [ctypes.c_char_p, ctypes.c_char_p]
    lib.TIFFOpen.restype = ctypes.c_void_p
    lib.TIFFReadEncodedStrip.argtypes = [ctypes.c_void_p, ctypes.c_uint32,
                                       ctypes.c_void_p, ctypes.c_ssize_t]
    lib.TIFFReadEncodedStrip.restype = ctypes.c_ssize_t
    lib.TIFFClose.argtypes = [ctypes.c_void_p]
    handle = lib.TIFFOpen(str(SOURCE).encode(), b'r')
    if not handle:
        raise RuntimeError('Cannot open source GeoTIFF')
    band = np.empty((TILE, width, 3), dtype=np.uint8)
    buffer = ctypes.create_string_buffer(rows * width * 3)
    strip_array = np.frombuffer(buffer, dtype=np.uint8)
    level0 = OUT / '0'
    level0.mkdir(exist_ok=True)
    band_rows, band_y = 0, 0
    try:
        for strip in range(strip_count):
            count = lib.TIFFReadEncodedStrip(handle, strip, buffer, -1)
            if count <= 0 or count % (width * 3):
                raise RuntimeError('Cannot decode TIFF strip ' + str(strip))
            decoded = strip_array[:count].reshape((-1, width, 3))
            offset = 0
            while offset < len(decoded):
                take = min(TILE - band_rows, len(decoded) - offset)
                band[band_rows:band_rows+take] = decoded[offset:offset+take]
                band_rows += take
                offset += take
                if band_rows == TILE or (strip == strip_count-1 and offset == len(decoded)):
                    for tx, x in enumerate(range(0, width, TILE)):
                        tile = Image.fromarray(band[:band_rows, x:min(width, x+TILE)].copy())
                        tile.save(level0 / (str(tx) + '_' + str(band_y) + '.jpg'),
                                  quality=88, subsampling=0)
                    print('Map strips: {0}/{1} rows'.format(min(height,(band_y+1)*TILE),height), flush=True)
                    band_rows = 0
                    band_y += 1
    finally:
        lib.TIFFClose(handle)

    levels = [{'level': 0, 'width': width, 'height': height}]
    level, prev_w, prev_h = 0, width, height
    while max(prev_w, prev_h) > TILE:
        level += 1
        out = OUT / str(level)
        out.mkdir(exist_ok=True)
        new_w, new_h = math.ceil(prev_w/2), math.ceil(prev_h/2)
        for ty in range(math.ceil(new_h/TILE)):
            for tx in range(math.ceil(new_w/TILE)):
                x0, y0 = tx*TILE*2, ty*TILE*2
                pw, ph = min(TILE*2, prev_w-x0), min(TILE*2, prev_h-y0)
                parent = Image.new('RGB', (pw, ph))
                for cy in range(2):
                    for cx in range(2):
                        child = OUT / str(level-1) / (str(tx*2+cx)+'_'+str(ty*2+cy)+'.jpg')
                        if cx*TILE < pw and cy*TILE < ph:
                            with Image.open(child) as image:
                                parent.paste(image, (cx*TILE, cy*TILE))
                parent = parent.resize((math.ceil(pw/2),math.ceil(ph/2)), Image.LANCZOS)
                parent.save(out / (str(tx)+'_'+str(ty)+'.jpg'),quality=90,subsampling=0)
        levels.append({'level': level, 'width': new_w, 'height': new_h})
        prev_w, prev_h = new_w, new_h
        print('Map pyramid: level ' + str(level), flush=True)

    def longitude(x):
        return math.degrees(x / 6378137)
    def latitude(y):
        return math.degrees(math.atan(math.sinh(y / 6378137)))
    metadata = {'label': 'Геологическая карта', 'source': SOURCE.name, 'crs': 'EPSG:3857',
                'width': width, 'height': height, 'pixel_size': [dx, dy],
                'bounds_mercator': [west,south,east,north],
                'bounds': [longitude(west),latitude(south),longitude(east),latitude(north)],
                'tile_size': TILE, 'levels': levels, 'max_level': level,
                'template': 'map_overlay/{level}/{x}_{y}.jpg', 'source_stamp': stamp}
    metadata_path.write_text(json.dumps(metadata,ensure_ascii=False,indent=2), encoding='utf-8')
    return metadata


if __name__ == '__main__':
    print(json.dumps(build_map_overlay(),ensure_ascii=False,indent=2))
