.PHONY: all
all: addin

.PHONY: addin
addin: bin/CustomMacro.pptm

MODULES = Module1.vb
CLASSES = Class1.vb

bin/CustomMacro.pptm: $(MODULES) $(CLASSES)
	vbamc -m $(MODULES) -c $(CLASSES) -n "Sample Addin" --company "ACME"

.PHONY: clean
clean:
	rm -rf bin/ obj/
