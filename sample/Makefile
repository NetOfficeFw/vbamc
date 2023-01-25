# Copyright 2022 Cisco Systems, Inc.
# Licensed under MIT-style license (see LICENSE.txt file).

.PHONY: all
all: addin

.PHONY: addin
addin: bin/PresentationAddin.ppam

MODULES = Module1.vb
CLASSES = Class1.vb

bin/PresentationAddin.ppam: $(MODULES) $(CLASSES)
	../src/vbamc/bin/Debug/net7.0/vbamc -m $(MODULES) -c $(CLASSES) -n "Sample Addin" --company "ACME" --user-profile-path "/Users/demo"

.PHONY: clean
clean:
	rm -rf bin/ obj/
