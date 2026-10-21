# Quire VB6 engine. The bundle is gitignored.
#   make

.PHONY: all engine

all: engine

engine:
	tools/vb6/build-engine.sh
