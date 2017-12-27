package main

import (
	"archive/zip"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"math"
	"os"
	"regexp"
	"strings"

	"github.com/richardlehane/mscfb"
)

var (
	olefilepath  = ""
	selectItem   = 0
	fatal        = true
	notFatal     = false
	olefileMagic = []byte{'\xD0', '\xCF', '\x11', '\xE0'}
)

// VBA macro source code is compressed in a stream
func findCompression(data []byte) []int {
	regexAttr := regexp.MustCompile("\x00Attribut\x00e ")
	if foundIndex := regexAttr.FindIndex(data); foundIndex != nil {
		return foundIndex
	}
	return nil
}

func parseTokenSequence(data []byte) ([][]byte, []byte) {
	flags := int(data[0])
	data = data[1:]
	var result [][]byte
	masks := []int{0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80} // WTF?

	for _, mask := range masks {
		if len(data) > 0 {
			if flags&mask != 0 {
				result = append(result, data[0:2])
				data = data[2:]
			} else {
				result = append(result, data[0:1])
				data = data[1:]
			}
		}
	}
	return result, data
}

func offsetBits(data []byte) int {
	numberOfBits := int(math.Ceil(math.Log2(float64(len(data)))))
	if numberOfBits < 4 {
		numberOfBits = 4
	} else if numberOfBits > 12 {
		numberOfBits = 12
	}

	return numberOfBits
}

func decompressChunk(compressedChunk []byte) ([]byte, []byte) {
	if len(compressedChunk) < 2 {
		return nil, nil
	}

	header := int(compressedChunk[0]) + (int(compressedChunk[1]))*0x100 // Macro = 01 10B2 => put the header in the correct order => B210
	size := (header & 0xFFF) + 3                                        // WTF?
	fmt.Printf("size = %d\n", size)
	flagCompressed := header & 0x8000 // WTF?
	fmt.Printf("%x\n", flagCompressed)
	data := compressedChunk[2 : 2+size-2]

	if flagCompressed == 0 {
		return data, compressedChunk[size:]
	}

	decompressedChunk := ""
	lengthData := len(data)
	var tokens [][]byte

	for lengthData != 0 {
		tokens, data = parseTokenSequence(data)
		lengthData = len(data)
		for _, token := range tokens {
			if len(token) == 1 {
				decompressedChunk += string(token)
			} else {
				if decompressedChunk == "" {
					return nil, nil
				}

				numberOfOffsetBits := offsetBits([]byte(decompressedChunk))
				copyToken := (int(token[0]) + int(token[1])*0x100)

				offset := 1 + (copyToken >> (uint(16 - numberOfOffsetBits)))
				length := 3 + (((copyToken << uint(numberOfOffsetBits)) & 0xFFFF) >> uint(numberOfOffsetBits))
				decompressedChunkTmp := []byte(decompressedChunk)
				copy := decompressedChunkTmp[len(decompressedChunkTmp)-offset:]
				copy = copy[0:length]
				lengthCopy := len(copy)

				for length > lengthCopy {
					if length-lengthCopy >= lengthCopy {
						for _, copyByte := range copy[0:lengthCopy] {
							copy = append(copy, copyByte)
						}
						length -= lengthCopy
					} else {
						for _, copyByte := range copy[0 : length-lengthCopy] {
							copy = append(copy, copyByte)
						}
						length -= length - lengthCopy
					}
				}
				decompressedChunk += string(copy)
			}
		}
	}

	return []byte(decompressedChunk), compressedChunk[size:]
}

func decompress(compressedData []byte) (bool, string) {
	if string(compressedData[0]) != string(1) {
		return false, ""
	}

	remainder := compressedData[1:]
	decompressed := ""
	lengthRemainder := len(remainder)
	var decompressedChunck []byte

	for lengthRemainder != 0 {
		decompressedChunck, remainder = decompressChunk(remainder)
		lengthRemainder = len(remainder)

		if decompressedChunck == nil {
			return false, decompressed
		}
		decompressed += string(decompressedChunck)
	}
	decompressed = strings.Replace(decompressed, "\r\n", "\n", -1)
	decompressed = strings.Replace(decompressed, "\x00\x00", "", -1)
	return true, decompressed
}

func checkError(err error, errAllowed bool) bool {
	if err != nil && errAllowed == notFatal {
		log.Println(err)
		return true
	} else if err != nil && errAllowed == fatal {
		log.Fatalln(err)
	}
	return false
}

func openFile() (*os.File, *zip.ReadCloser) {
	reader, err := zip.OpenReader(olefilepath)
	if err != nil {
		// we don't care if it's not zipped, not important
		file, err := os.Open(olefilepath)
		checkError(err, fatal)
		return file, nil
	}
	return nil, reader

}

func processOle(file *os.File) {
	doc, err := mscfb.New(file)
	resultErr := checkError(err, notFatal)
	if resultErr == true {
		return
	}

	entryNumber := 0

	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		entryNumber++
		foundCompr := ""
		buf := make([]byte, entry.Size)
		i, _ := doc.Read(buf)

		if i > 0 {
			indexCompressedData := findCompression(buf[:i])
			if indexCompressedData != nil {
				foundCompr = "M"
				if selectItem == entryNumber {
					result, data := decompress(buf[indexCompressedData[0]-3 : i])
					if result == true {
						fmt.Println(data)
					} else {
						fmt.Println("Could not decompress item selected")
						os.Exit(1)
					}
				}
			}
		}
		if selectItem == 0 {
			fmt.Println(entryNumber, "\t", foundCompr, "\t", entry.Size, "\t\t", entry.Name)
		}
	}
}

func run() {
	// TODO: Test zip file and extract useful content
	file, zFiles := openFile()
	if file == nil {
		for _, zFile := range zFiles.File {
			rc, err := zFile.Open()
			checkError(err, fatal)
			buf := make([]byte, zFile.UncompressedSize)
			rc.Read(buf)
			magic := buf[0:4]
			magicCounter := 0
			for index := range magic {
				if magic[index] == olefileMagic[index] {
					magicCounter++
				}
				if magicCounter == 4 {
					if len(strings.Split(zFile.Name, "/")) > 1 {
						errMkdir := os.MkdirAll("./"+zFile.Name, 0777)
						checkError(errMkdir, true)
					}
					tmpfile, err := ioutil.TempFile(".", zFile.Name)
					checkError(err, fatal)
					_, errTmp := tmpfile.Write(buf)
					checkError(errTmp, fatal)
					rc.Close()
					processOle(tmpfile)
					//os.Remove(tmpfile.Name())
					//os.RemoveAll(strings.Split(zFile.Name, "/")[0])
				}
			}
		}

	} else {
		processOle(file)
		file.Close()
	}

}

func usage() {
	message := ""
	message += "missing arguments.\n"
	message += "Example of usage:\n"
	message += "goledump --olefilepath /path/to/olefile\t\t\t\t\t Show different items in the document\n"
	message += "goledump --olefilepath /path/to/olefile --select 8\t\t\t Dump the element number 8"
	fmt.Println(message)
	os.Exit(0)
}

func init() {
	flag.StringVar(&olefilepath, "olefilepath", "", "OLE file path")
	flag.IntVar(&selectItem, "select", 0, "select item number for dumping")
}

func main() {
	if len(os.Args) < 2 {
		usage()
	}
	flag.Parse()
	run()
}
