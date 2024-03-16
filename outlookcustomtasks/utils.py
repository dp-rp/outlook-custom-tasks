# NOTE: only for assisting with development
def _inspect_thing(thing_name,thing):
    print("----------")
    print(thing_name)
    print(type(thing))
    print(dir(thing))
    for key in dir(thing):
        try:
            print(key)
            print(getattr(thing,key))
            print()
        except KeyboardInterrupt as err:
            raise err
        except Exception as err:
            print(f"Error: {key}: Failed to get value: {err}")
    print("----------")
