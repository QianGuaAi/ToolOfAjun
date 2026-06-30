This folder contains optional native binary placeholders that ship next to MyTools.exe.

Layout:
NativeBinaries\ffmpeg\README.txt   (FFmpeg is external; install it separately or copy ffmpeg.exe here after build/install)

Files like frpc.exe.gz are embedded into MyTools.exe at build time and do not appear here in the published output.
The build and installer intentionally exclude NativeBinaries\ffmpeg\ffmpeg.exe.
