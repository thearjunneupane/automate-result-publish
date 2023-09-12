DEMO_ENTRY=demo_entry/demo_entry.exe
PUBLISH=publish_result/publish.exe
BINARY_NAME=main.exe

build:
	go build -o ${DEMO_ENTRY} demo_entry/demo_entry.go
	go build -o ${PUBLISH} publish_result/publish.go
	go build -o ${BINARY_NAME} main.go

run: build
	./${DEMO_ENTRY}
	./${PUBLISH}
	./${BINARY_NAME}

clean:
	go clean
	rm -f ${DEMO_ENTRY} ${PUBLISH} ${BINARY_NAME}
