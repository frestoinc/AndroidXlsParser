import LocaleHeader.Companion.toFolderName
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import kotlin.io.path.Path
import kotlin.io.path.createDirectory


enum class LocaleHeader(
    val customDisplayName: String,
    val code: String
) {
    ENGLISH("English", "en"),
    SIMPLIFIED_CHINESE_CHINA("Simplified Chinese", "zh-CN"),
    TRADITIONAL_CHINESE_CHINA("Trad Chinese", "zh-TW"),
    JAPANESE("Japanese", "ja-JP"),
    KOREAN("Korean", "ko-KR"),
    MALAY("Malay", "ms-MY");

    companion object {
        fun String.toFolderName(): String {
            val header = values().find { it.customDisplayName == this } ?: ENGLISH
            return "values-b+${header.code}"
        }
    }
}

fun main(args: Array<String>) {
    val sheet = getSheet("sample/translations.xlsx", 0)
    val maps = generateMapValues(sheet)
    convertMapToXml(maps)
}

/**
 * Return instance of [XSSFSheet] based on the [index] of the File[filename]
 * @param filename name of the File
 * @param index index of spreadsheet
 */
@Suppress("SameParameterValue")
private fun getSheet(filename: String, index: Int): XSSFSheet {
    val inputStream = FileInputStream(filename)
    val xssWb = XSSFWorkbook(inputStream)
    return xssWb.getSheetAt(index)
}

/**
 * Return the list of column 0
 * @param sheet of the [XSSFWorkbook]
 */
private fun getKeysWithinXlsMap(sheet: XSSFSheet): List<String> =
    sheet.mapNotNull { it.getCell(0) }.mapNotNull { it.stringCellValue }.toList()

/**
 * Return a [Map] instance of the [sheet]
 *
 * @param sheet of the [XSSFWorkbook]
 */
private fun generateMapValues(sheet: XSSFSheet): Map<String, List<String>> =
    getKeysWithinXlsMap(sheet).mapIndexed { index, keys ->
        keys to sheet.getRow(index)
    }.associate { rows ->
        rows.first to rows.second.mapIndexedNotNull { index, _ -> rows.second.getCell(index + 1) }
            .mapNotNull { it.stringCellValue }
    }

/**
 * Take the first sheet row and return a map against [LocaleHeader]
 *
 * Generate folders based on the return map
 *
 * Iterate through the map and write data retrieved from [maps]
 *
 * @param maps sheet of the [XSSFWorkbook]
 */
private fun convertMapToXml(maps: Map<String, List<String>>) {
    val headers = maps.toList().take(1).associate { pair ->
        pair.first to pair.second.map { it.toFolderName() }
    }.values.flatten()
    println(headers)
    headers.map { it.createPaths() }

    headers.mapIndexed { index, s ->
        val sb = StringBuilder()
        sb.append(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
                    "<resources>\n"
        )
        try {
            maps.toList().drop(1).map { pair ->
                sb.append("\t<string name=${pair.first}>${pair.second[index]}</string>\n")
            }
        } catch (e: Exception) {
            return@mapIndexed
        }

        sb.append("</resources>")
        println(sb)
        writeToFile(s, data = sb.toString())
    }
}

private fun String.createPaths() {
    val path = Path(this)
    File(this).deleteRecursively()
    path.createDirectory()
}

private fun writeToFile(folderName: String, data: String) {
    val file = File(folderName, "strings.xml")
    if (file.exists()) {
        file.delete()
    }
    file.createNewFile()
    file.writeText(data)
}