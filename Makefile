clean:
	rm -r certificates
	rm -r ppt

install:
	pip3 install -r requirements.txt

lint:
	pylint certify.py GoogleApi/gmail_api.py

run :
	python3 main.py

sandbox:
	