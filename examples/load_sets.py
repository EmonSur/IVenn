from ivenn import IVenn

v = IVenn.from_excel("datasets/6set_dataset.xlsx")

v.set_unions("(((A,B),E),(C,D))")

v.draw()