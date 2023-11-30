#!/bin/bash

argVersion="$1"

version=${argVersion:-0.0.1}
configuration=${configuration:-Debug}
outputdir=${outputdir:-$(cygpath -w $HOME/nugetpacks)}

#echo Outputdir is: $outputdir
cmd /c $(echo $APPDATA/bin/nuget pack RpaLib.csproj -Version $version -OutputDirectory "$outputdir" -Properties "Configuration=$configuration")