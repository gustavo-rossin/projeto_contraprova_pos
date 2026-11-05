"""Converte o GeoJSON de UFs em paths SVG simplificados (Douglas-Peucker)."""
import json, math, sys
g = json.load(open(sys.argv[1]))
K = math.cos(math.radians(-15)); S = 20  # px por grau
LON0, LAT0 = -74.2, 5.5
def proj(p): return ((p[0]-LON0)*K*S, (LAT0-p[1])*S)
def dp(pts, eps):
    if len(pts) < 3: return pts
    (x1,y1),(x2,y2) = pts[0], pts[-1]; dx,dy = x2-x1, y2-y1; L = math.hypot(dx,dy) or 1e-9
    i,dm = 0,0
    for k in range(1,len(pts)-1):
        d = abs(dy*pts[k][0]-dx*pts[k][1]+x2*y1-y2*x1)/L
        if d > dm: i,dm = k,d
    if dm > eps: return dp(pts[:i+1],eps)[:-1] + dp(pts[i:],eps)
    return [pts[0], pts[-1]]
def area(r): return abs(sum(r[i][0]*r[i-1][1]-r[i-1][0]*r[i][1] for i in range(len(r))))/2
out = {}; W = H = 0
for f in g["features"]:
    uf = f["properties"]["sigla"]; geom = f["geometry"]
    polys = geom["coordinates"] if geom["type"] == "MultiPolygon" else [geom["coordinates"]]
    d = []; cx = cy = A = 0
    for poly in polys:
        ring = [proj(p) for p in poly[0]]
        a = area(ring)
        if a < 2: continue
        m = len(ring)//2  # anel fechado: simplifica em duas metades
        r = dp(ring[:m+1], 0.6)[:-1] + dp(ring[m:], 0.6)
        if len(r) < 4: continue
        d.append("M" + "L".join(f"{x:.1f},{y:.1f}" for x,y in r) + "Z")
        if a > A: A = a; cx = sum(p[0] for p in ring)/len(ring); cy = sum(p[1] for p in ring)/len(ring)
        W = max(W, max(p[0] for p in r)); H = max(H, max(p[1] for p in r))
    out[uf] = {"d": "".join(d), "cx": round(cx,1), "cy": round(cy,1)}
json.dump({"w": math.ceil(W)+4, "h": math.ceil(H)+4, "ufs": out}, open(sys.argv[2], "w"))
print(W, H, sum(len(v["d"]) for v in out.values()))
