package main

import (
	"flag"
	"fmt"
	"log"
	"math"
	"os"
	"regexp"
	"strings"

	"github.com/richardlehane/mscfb"
)

var (
	olefilepath = ""
	selectItem  = ""
)

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

	header := int(compressedChunk[0]) + (int(compressedChunk[1]))*0x100 // WFT?
	size := (header & 0xFFF) + 3                                        // WTF?
	flagCompressed := header & 0x8000                                   // WTF?
	data := compressedChunk[2 : 2+size-2]
	fmt.Println(len(data))
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
	return true, strings.Replace(decompressed, "\r\n", "\n", -1)
}

func searchAndDecompressSub(data []byte) (bool, string) {
	position := findCompression(data)
	if position == nil {
		return false, ""
	}
	compressedData := data[position[0]-3:]
	return decompress(compressedData)
}

func searchAndDecompress(data []byte) string {
	result, decompress := searchAndDecompressSub(data)
	if result == true {
		return decompress
	} else {
		return ""
	}
}

func checkError(err error) {
	if err != nil {
		log.Fatal(err)
	}
}

func run() {
	file, err := os.Open(olefilepath)
	checkError(err)

	defer file.Close()

	doc, err := mscfb.New(file)
	checkError(err)

	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		foundCompr := ""
		buf := make([]byte, entry.Size)
		i, _ := doc.Read(buf)

		if i > 0 {
			indexCompressedData := findCompression(buf[:i])
			if indexCompressedData != nil {
				foundCompr = "M"
				//_, result := decompress(buf[indexCompressedData[0]-3 : i])
				//fmt.Println(result)
			}
			if selectItem == "" {
				fmt.Println(foundCompr+"\t | \t", entry.Name)
			}
		}
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
	flag.StringVar(&selectItem, "select", "", "select item number for dumping")
}

func main() {
	if len(os.Args) < 2 {
		usage()
	}
	flag.Parse()
	run()
}
