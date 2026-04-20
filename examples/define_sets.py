from ivenn import IVenn, Set

a = Set("A", [1,3,6,7,8])
b = Set("B", [1,3,6,7,8])
c = Set("C", [1,4,7,8])
d = Set("D", [2,4,5,6,7])
e = Set("E", [1,17,18,19,20])
f = Set("F", [5,6,8,21,22,23])

v = IVenn(a, b, c, d, e, f)
v.set_unions("ab")
v.draw()