# Final Solution Directory

**This is the final solution directory where the Solution.zip and Solution_managed.zip files are stored.**

Navigate to this folder in the terminal and run the following commands.
```
$ pac solution init --publisher-name Bever --publisher-prefix bvr
$ pac solution add-reference --path ../
$ msbuild /t:rebuild /restore /p:configuration=release
```

> ***Note*** The msbuild command must be run in the Developer Command Prompt.

The generated solution files are located inside the \bin\Release\ folder after the build is successful.
