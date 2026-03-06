
.PHONY: docxtomd classify rezult build clean deletecache

unzip:
	python unzip_in_folder.py ..

docxtomd:
	python docx_to_markdown.py ..

classify:
	cd .. && python tools/classify_md_by_class.py --move

rezult:
	python copy_data_to_file.py

deletecache :
	rm -rf 团建联络员工作记录表.md

build: unzip docxtomd classify deletecache

clean :
	cd .. && rm -rf *.md && find . -type d -name "本*" -exec rm -rf {} +
	rm -rf 团建联络员工作记录表.md

# read and excute step 1 2 3