import sys, json
x = sys.argv[1]
sys.path.insert(0, '/mnt/user-data/uploads/Downloads/Cantina-app/lib')
src = open('/mnt/user-data/uploads/Downloads/Cantina-app/server.py').read().split('# ---------- HTTP')[0]
sys.argv = ['server.py', x]
g = {'__file__': '/mnt/user-data/uploads/Downloads/Cantina-app/server.py', '__name__': 'srv'}
exec(src, g)
d = g['read_all'](g['openpyxl'].load_workbook(x))
print(json.dumps({k: [[v for kk, v in o.items() if kk != 'row'] for o in vv] for k, vv in d.items()}, default=str))
