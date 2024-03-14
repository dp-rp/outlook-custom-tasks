# NOTE: only for assisting with development
def _inspect_thing(thing_name,thing):
    print("----------")
    print(thing_name)
    print(type(thing))
    print(dir(thing))
    for key in dir(thing):
        print(key)
        print(getattr(thing,key))
        print()
    print("----------")
