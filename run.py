import importlib

def _load_target():
    module = importlib.import_module('word_template_pro.__main__')
    obj = module
    for part in 'main'.split("."):
        obj = getattr(obj, part)
    return obj

def main():
    target = _load_target()
    result = target()
    raise SystemExit(result if isinstance(result, int) else 0)

if __name__ == "__main__":
    main()
