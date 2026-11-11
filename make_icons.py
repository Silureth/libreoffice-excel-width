"""Generate the toolbar icons (no dependencies): |<-->| on transparent bg."""
import struct, zlib, sys, os

def png(path, n):
    col = (0x25, 0x63, 0xeb, 255)                       # blue, ok on light & dark
    px = [[(0, 0, 0, 0)] * n for _ in range(n)]
    def put(x, y):
        if 0 <= x < n and 0 <= y < n: px[y][x] = col
    t = max(1, n // 13)                                  # stroke thickness
    m = n // 2
    for y in range(n // 6, n - n // 6):                  # column borders
        for k in range(t): put(1 + k, y); put(n - 2 - k, y)
    for x in range(2 + t, n - 2 - t):                    # shaft
        for k in range(t): put(x, m - t // 2 + k)
    h = n // 4                                           # arrow heads
    for i in range(h):
        for k in range(t):
            for y in (m - i, m + i):
                put(2 + t + i + k, y); put(n - 3 - t - i - k, y)
    raw = b"".join(b"\0" + bytes(c for p in row for c in p) for row in px)
    chunk = lambda tag, data: struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xffffffff)
    with open(path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", n, n, 8, 6, 0, 0, 0))
                + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))

out = sys.argv[1]
png(os.path.join(out, "width_16.png"), 16)
png(os.path.join(out, "width_26.png"), 26)
