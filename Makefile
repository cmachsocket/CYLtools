
.PHONY: docxtomd classify rezult build clean deletecache

docxtomd:
	python docx_to_markdown.py ..

classify:
	cd .. && python tools/classify_md_by_class.py --move

rezult:
	python csv_to_excel.py --input 团建联络员打分结果_自动评估.csv --overwrite && \
	python csv_to_excel.py --input 团建联络员工作记录表_填写结果.csv --overwrite 

deletecache :
	rm -rf 团建联络员工作记录表陈明远.md

build: docxtomd classify deletecache

clean :
	cd .. && rm -rf *.md && find . -type d -name "本*" -exec rm -rf {} +
	rm -rf 团建联络员工作记录表_填写结果.xlsx 
	rm -rf 团建联络员打分结果_自动评估.xlsx
	rm -rf 团建联络员工作记录表陈明远.md

# read and excute step 1 2 3