# Copyright 2022 Cisco Systems, Inc.
# Licensed under MIT-style license (see LICENSE.txt file).

framework := net10.0

.PHONY: all
all: addin

.PHONY: addin
addin: bin/PresentationAddin.ppam bin/PresentationAddin.pptm

MODULES = Module1.vb
CLASSES = Class1.vb

bin/PresentationAddin.ppam: $(MODULES) $(CLASSES)
	../src/vbamc/bin/Debug/$(framework)/vbamc -m $(MODULES) -c $(CLASSES) -f "PresentationAddin.ppam" -n "Sample Addin" --company "ACME" -p Generator=VbaCompiler

bin/PresentationAddin.pptm: $(MODULES) $(CLASSES)
	../src/vbamc/bin/Debug/$(framework)/vbamc -m $(MODULES) -c $(CLASSES) -f "PresentationAddin.pptm" -n "Sample Addin" --company "ACME" -p Generator=VbaCompiler

.PHONY: clean
clean:
	rm -rf bin/ obj/
