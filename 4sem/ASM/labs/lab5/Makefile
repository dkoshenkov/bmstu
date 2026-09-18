SRC_DIR = src
BUILD_DIR = build

TARGET = main
ASM_MODULE = change_order_asm
EXTRA_MODULE = out

all: run

build: ${ASM_MODULE} ${TARGET} ${EXTRA_MODULE}
	g++ -z noexecstack -no-pie ${BUILD_DIR}/${ASM_MODULE}.o ${BUILD_DIR}/${TARGET}.o ${BUILD_DIR}/${EXTRA_MODULE}.o -o ${BUILD_DIR}/${TARGET}.out

${ASM_MODULE}: ${SRC_DIR}/${ASM_MODULE}.asm	
	nasm -f elf64 ${SRC_DIR}/${ASM_MODULE}.asm -o ${BUILD_DIR}/${ASM_MODULE}.o

${TARGET}: ${SRC_DIR}/${TARGET}.cpp
	g++ -std=c++20 -c ${SRC_DIR}/${TARGET}.cpp -o ${BUILD_DIR}/${TARGET}.o

${EXTRA_MODULE}: ${SRC_DIR}/${EXTRA_MODULE}.cpp
	g++ -std=c++20 -c ${SRC_DIR}/${EXTRA_MODULE}.cpp -o ${BUILD_DIR}/${EXTRA_MODULE}.o

run: build
	./${BUILD_DIR}/${TARGET}.out

debug: build
	edb --run ${BUILD_DIR}/${TARGET}.out