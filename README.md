EasyAddin for Excel is an add-in with some useful functions and functionality.

# Features

EasyAddin provides the following features:

- 1D interpoloation

# Getting Started

## Usage of an add-in

Get EasyAddin from the latest release and install it as any other add-in in Excel.

## For developers

EasyAddin has a version control through `git hooks` written in Python. More about git hooks read [here](https://www.xltrail.com/blog/auto-export-vba-commit-hook).

In order to automagically export workbook’s VBA modules into stand-alone text files on every git commit, you need:

- Python with `oletools` installed and
- the files `pre-commit` and `pre-commit.py` in the `.git/hooks` directory

Nowadays, it is possible to keep `git hooks` under version control as the original `.git/hooks` folder is not tracked. The hooks for this project are stored in `.hooks` folder and you need to config `.hooks` directory as the core path for hooks, e.g., MY_REPO_DIR/hooks would be

```
git config --local core.hooksPath hooks/
```

With that in place, any git commit automatically takes care of dumping your workbooks VBA content as text files to your filesystem. Which is something that Git understands well.
