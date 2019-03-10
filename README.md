# Link Utilities #
A utility for working with Windows shortcut (.lnk) files

## Usage ##

```
Usage: LnkUtils (command) [options]
 Commands:

  Check (path to folder or .lnk file) [options]
  - tests link files for bad paths

  Check Options:
  -r        recurse into sub-folders
  -v        print extra information about links
  -a        show all links not just bad ones

  Create [options]
  - create a new link

  Create Options:
  -t (path)          path to target
  -n (path)          name of link file (defaults to 'target file name'.lnk)
  -c (text)          comment
  -s (path)          start in path
  -i (path) [index]  path to .ico file or .dll/.exe. index defaults to 0
  -f                 force create even if target doesn't exist
```

## Notes ##
I came accross this gem while searching for information oniline - [MS-SHLLINK: Shell Link \(.LNK\) Binary File Format](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-shllink/16cb4ca1-9339-4d0c-a68d-bf1d6cc0f943).

Maybe a future version of this tool could expose a way to create lnk files directly using the published format instead of the win32 api.