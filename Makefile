TARGET ?= /mnt/c/tools
BINDIR ?= $(HOME)/.local/bin

.PHONY: install clean

install:
	@mkdir -p "$(TARGET)"
	@cp -f pptx2jpg.py "$(TARGET)/pptx2jpg.py"
	@echo "Installed $(TARGET)/pptx2jpg.py"
	@mkdir -p "$(BINDIR)"
	@cp -f pptx2jpg.sh "$(BINDIR)/pptx2jpg"
	@chmod +x "$(BINDIR)/pptx2jpg"
	@echo "Installed $(BINDIR)/pptx2jpg"

clean:
	@if [ -f "$(TARGET)/pptx2jpg.py" ]; then \
		rm -f "$(TARGET)/pptx2jpg.py"; \
		echo "Removed $(TARGET)/pptx2jpg.py"; \
	else \
		echo "$(TARGET)/pptx2jpg.py not found, nothing to remove"; \
	fi
	@if [ -f "$(BINDIR)/pptx2jpg" ]; then \
		rm -f "$(BINDIR)/pptx2jpg"; \
		echo "Removed $(BINDIR)/pptx2jpg"; \
	else \
		echo "$(BINDIR)/pptx2jpg not found, nothing to remove"; \
	fi
