$python_script_path = python .\.util-scripts\get_python_script_path.py  # get path to Python user scripts
pushd $python_script_path  # temporarily move to dir with Python user scripts
.\pipx.exe ensurepath  # run `pipx ensurepath` to set up pipx PATH requirements
popd  # move back to original directory