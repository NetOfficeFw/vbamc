# Copyright 2022 Cisco Systems, Inc.
# Licensed under MIT-style license (see LICENSE.txt file).

.PHONY: all
all: addin

.PHONY: addin
addin: bin/CustomMacro.pptm

MODULES = Module1.vb
CLASSES = Class1.vb

bin/CustomMacro.pptm: $(MODULES) $(CLASSES)
	d:\dev\github\NetOfficeFw\vbamc\src\vbamc\bin\Debug\net7.0\vbamc.exe -m $(MODULES) -c $(CLASSES) -n "Sample Addin" --company "ACME"

.PHONY: clean
clean:
	rm -rf bin/ obj/
