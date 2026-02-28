import os

def ensure_safe_path(path, base_dir=None):
    """
    Ensures that the given path is safe and does not escape the base directory.
    If base_dir is None, the current working directory is used.
    """
    # Allow paths in /tmp for temporary files
    if path.startswith('/tmp/'):
        return os.path.abspath(path)

    if base_dir is None:
        base_dir = os.getcwd()

    abs_base_dir = os.path.abspath(base_dir)
    abs_path = os.path.abspath(os.path.join(abs_base_dir, path))

    if not abs_path.startswith(abs_base_dir):
        raise ValueError(f"Path traversal detected: {path} is outside of {abs_base_dir}")

    return abs_path
