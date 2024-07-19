package xyz.proteanbear.template;

import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import xyz.proteanbear.template.annotation.PbPOIWordVariable;
import xyz.proteanbear.template.exception.FileSuffixNotSupportException;
import xyz.proteanbear.template.utils.ClassUtils;

import java.io.*;
import java.lang.annotation.Annotation;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.URL;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * The tools is created for easy use of Apache POI.
 * This tool is used for reading and writing in Word.
 *
 * @author ProteanBear
 */
public class PbPOIWordTemplate
{
    private static final Logger logger = LoggerFactory.getLogger(PbPOIWordTemplate.class);

    /**
     * Record image insertion related content
     */
    public static class Image
    {
        private String url;
        private final PictureType type;
        private final String description;
        private final int width;
        private final int height;

        public Image(String url, PictureType type, String description, int width, int height)
        {
            this.url = url;
            this.type = type;
            this.description = description;
            this.width = width;
            this.height = height;
        }

        public String getUrl()
        {
            return url;
        }

        public void setUrl(String url)
        {
            this.url = url;
        }

        public PictureType getType()
        {
            return type;
        }

        public String getDescription()
        {
            return description;
        }

        public int getWidth()
        {
            return width;
        }

        public int getHeight()
        {
            return height;
        }
    }

    /**
     * The start identifier of the replacement variable
     */
    private String variableStart = "{";

    /**
     * The end identifier of the replacement variable
     */
    private String variableEnd = "}";

    /**
     * The start identifier of the replacement table row loop data
     */
    private String loopStart = "[";

    /**
     * The end identifier of the replacement table row loop data
     */
    private String loopEnd = "]";

    /**
     * Set the start identifier of the replacement variable
     *
     * @param variableStart The start identifier of the replacement variable
     * @return current object
     */
    public PbPOIWordTemplate setVariableStart(String variableStart)
    {
        this.variableStart = variableStart;
        return this;
    }

    /**
     * Set the end identifier of the replacement variable
     *
     * @param variableEnd The start identifier of the replacement variable
     * @return current object
     */
    public PbPOIWordTemplate setVariableEnd(String variableEnd)
    {
        this.variableEnd = variableEnd;
        return this;
    }

    /**
     * Set the start identifier of the replacement table row loop data
     *
     * @param loopStart The start identifier of the replacement table row loop data
     * @return current object
     */
    public PbPOIWordTemplate setLoopStart(String loopStart)
    {
        this.loopStart = loopStart;
        return this;
    }

    /**
     * Set the end identifier of the replacement table row loop data
     *
     * @param loopEnd The start identifier of the replacement table row loop data
     * @return current object
     */
    public PbPOIWordTemplate setLoopEnd(String loopEnd)
    {
        this.loopEnd = loopEnd;
        return this;
    }

    /**
     * After reading the specified WORD template and replacing the contents of the template, a new WORD file is generated
     *
     * @param templateFile file must exist
     * @param toFile       new file for generated
     * @param data         Replace data objects with template variables declared by annotations
     * @throws IOException                   io exception
     * @throws NoSuchMethodException         No such method
     * @throws InvocationTargetException     Invocation target
     * @throws IllegalAccessException        Illegal access
     * @throws FileSuffixNotSupportException file suffix is not supported
     */
    public void writeTo(File templateFile, File toFile, Object data)
            throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException,
            FileSuffixNotSupportException
    {
        //File exists, delete the old file
        if (toFile.exists()) toFile.delete();
        toFile.createNewFile();

        writeTo(
                templateFile,
                Files.newOutputStream(toFile.toPath()),
                data
        );
    }

    /**
     * After reading the specified WORD template and replacing the contents of the template, a new WORD file is generated
     *
     * @param templateFile file must exist
     * @param outputStream Output stream
     * @param data         Replace data objects with template variables declared by annotations
     * @throws IOException               io exception
     * @throws NoSuchMethodException     No such method
     * @throws InvocationTargetException Invocation target
     * @throws IllegalAccessException    Illegal access
     */
    public void writeTo(File templateFile, OutputStream outputStream, Object data)
            throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException
    {
        //Read Word template file
        if (!templateFile.exists()) throw new IOException("Template file is not exist");

        //Convert data objects to mapping dictionaries based on annotations
        Map<String, Object> dataMap = ClassUtils.dataMapBy(PbPOIWordVariable.class, data, new ClassUtils.Wrapper()
        {
            @Override
            public <T extends Annotation> Object wrap(T annotation, Object data)
            {
                PbPOIWordVariable variable = (PbPOIWordVariable) annotation;
                //Image path
                return variable.isImagePath()
                       ? new PbPOIWordTemplate.Image(
                        String.valueOf(data),
                        variable.imageType(),
                        variable.imageDescription(),
                        variable.imageWidth(),
                        variable.imageHeight()
                ) : data;
            }
        });

        try (XWPFDocument template = new XWPFDocument(Files.newInputStream(templateFile.toPath())))
        {
            //Get ordinary paragraphs and replace all variables
            List<XWPFParagraph> paragraphList = template.getParagraphs();
            //Replace variables in the content
            paragraphList.forEach(paragraph -> replaceParagraph(paragraph, dataMap));

            //Get table paragraphs and replace
            List<XWPFTable> tables = template.getTables();
            //Replace variables in the tables
            tables.forEach(table -> replaceTableParagraph(table, dataMap));

            //Write Word file
            template.write(outputStream);
        }
    }

    /**
     * Replace the corresponding variable in the field content
     *
     * @param paragraph include a lot of runs
     * @param dataMap   data map
     */
    private void replaceParagraph(XWPFParagraph paragraph, Map<String, Object> dataMap)
    {
        //Search variable
        TextSegment start = paragraph.searchText(variableStart, new PositionInParagraph());
        if (start == null) return;
        TextSegment end = paragraph.searchText(
                variableEnd,
                new PositionInParagraph(start.getEndRun(), start.getEndText(), start.getEndChar())
        );
        if (end == null) return;

        //Find the replaced variable and tag
        String variableKey = paragraph.getText(new TextSegment(
                start.getBeginRun(), end.getEndRun(),
                start.getBeginText(), end.getEndText(),
                start.getBeginChar(), end.getEndChar()
        ));
        variableKey = variableKey.replace(variableStart, "")
                                 .replace(variableEnd, "");
        Object variable = dataMap.getOrDefault(variableKey, "");

        //Image
        if (variable instanceof Image)
        {
            replacePictureParagraph(paragraph, start, end, (Image) variable, variableEnd.length());
        }
        else
        {
            String variableValue = (variable == null ? "" : String.valueOf(variable));
            //Traversing Runs, replacing the contents of the variable
            replaceParagraph(paragraph, start, end, variableValue, variableEnd.length());
        }

        //Check again if there are any variables in the paragraph to replace
        replaceParagraph(paragraph, dataMap);
    }

    /**
     * Replace the corresponding variable in the field content
     *
     * @param paragraph         include a lot of runs
     * @param start             Replace the starting fragment position
     * @param end               Replace the ending fragment position
     * @param variableValue     Replaced content
     * @param variableEndLength Replacement terminator length
     */
    private void replaceParagraph(
            XWPFParagraph paragraph, TextSegment start, TextSegment end
            , String variableValue, int variableEndLength
    )
    {
        //Traversing Runs, replacing the contents of the variable
        List<XWPFRun> runs = paragraph.getRuns();
        XWPFRun runStart = runs.get(start.getBeginRun());
        XWPFRun runEnd = runs.get(end.getEndRun());
        String textBefore = runStart.text()
                                    .substring(0, start.getBeginChar());
        String textAfter = runEnd.text()
                                 .substring(end.getEndChar() + variableEndLength);
        StringBuilder builder = new StringBuilder();
        //If start-run and end-run is in the same sun
        if (start.getBeginRun() == end.getEndRun())
        {
            runStart.setText(
                    builder.append(textBefore)
                           .append(variableValue)
                           .append(textAfter)
                           .toString()
                    , 0);
        }
        //If start-run and end-run is in the different suns
        //Insert a new run after start-run
        else
        {
            runStart.setText(
                    builder.append(textBefore)
                           .append(variableValue)
                           .toString()
                    , 0);
            runEnd.setText(textAfter, 0);
            for (int pos = end.getEndRun(); pos > start.getBeginRun(); pos--)
            {
                paragraph.removeRun(pos);
            }
        }
    }

    /**
     * Replace the corresponding variable in the field content with the picture
     *
     * @param paragraph         include a lot of runs
     * @param start             Replace the starting fragment position
     * @param end               Replace the ending fragment position
     * @param image             Replaced content
     * @param variableEndLength Replacement terminator length
     */
    private void replacePictureParagraph(
            XWPFParagraph paragraph, TextSegment start, TextSegment end
            , Image image, int variableEndLength
    )
    {
        //Traversing Runs, replacing the contents of the variable
        List<XWPFRun> runs = paragraph.getRuns();
        XWPFRun runStart = runs.get(start.getBeginRun());
        XWPFRun runEnd = runs.get(end.getEndRun());
        String textAfter = runEnd.text()
                                 .substring(end.getEndChar() + variableEndLength);

        //Read the image file form the url
        try (InputStream imageInput = new URL(image.url).openStream())
        {
            runStart.addPicture(
                    imageInput,
                    image.type,
                    image.description,
                    Units.toEMU(image.width),
                    Units.toEMU(image.height)
            );
        }
        catch (IOException | InvalidFormatException e)
        {
            logger.error("Add a picture(url:{}) failed:", image.url, e);
        }

        //If start-run and end-run is in the same sun
        if (start.getBeginRun() == end.getEndRun())
        {
            runStart.setText(" ", 0);
        }
        //If start-run and end-run is in the different suns
        //Insert a new run after start-run
        else
        {
            runStart.setText(" ", 0);
            runEnd.setText(textAfter, 0);
            for (int pos = end.getEndRun(); pos > start.getBeginRun(); pos--)
            {
                paragraph.removeRun(pos);
            }
        }
    }

    /**
     * Replace the corresponding variable in the table
     *
     * @param table   include table paragraph
     * @param dataMap data map
     */
    private void replaceTableParagraph(XWPFTable table, Map<String, Object> dataMap)
    {
        //Record the rows used for data loops
        Map<Integer, XWPFTableRow> rowForLoop = new HashMap<>();
        //Record a description of the relevant substitution variables used for the data loop
        // (including start position, end position, variable index key name, and get method)
        Map<Integer, Map<Integer, List<Replacement>>> replace = new HashMap<>();
        //Method for obtaining corresponding fields in cached data
        Map<String, Map<String, Method>> methodCache = new HashMap<>();
        //Record the number of rows of data
        AtomicInteger dataLength = new AtomicInteger(0);

        //Traversing a table to replace normal variable content
        //And record and judge the rows that need data loops
        int[] index = {0, 0, 0};
        table.getRows()
             .forEach(row -> {
                 index[1] = 0;
                 row.getTableCells()
                    .forEach(cell -> {
                        //Check if the contents of the data loop are included in the cell
                        String content = cell.getText();
                        if (content.contains(loopStart)
                                && content.contains(loopEnd)
                                && !rowForLoop.containsKey(index[0]))
                        {
                            rowForLoop.put(index[0], row);
                        }

                        //Traversing paragraphs to replace common variables
                        //If a data loop is included, the search records all of the loop data descriptions in each cell
                        Map<Integer, List<Replacement>> map = replace.getOrDefault(index[1], new HashMap<>());
                        replace.put(index[1], map);
                        index[2] = 0;
                        cell.getParagraphs()
                            .forEach(paragraph -> {
                                replaceParagraph(paragraph, dataMap);
                                if (rowForLoop.containsKey(index[0]))
                                {
                                    List<Replacement> list = map.getOrDefault(index[2], new ArrayList<>());
                                    map.put(index[2]++, list);
                                    searchLoopVariable(
                                            paragraph,
                                            list,
                                            dataMap,
                                            methodCache,
                                            dataLength,
                                            new PositionInParagraph()
                                    );
                                }
                            });

                        index[1]++;
                    });
                 index[0]++;
             });

        //Traversing the rows of data that need to be looped,
        //inserting the looped rows of data under the rows of data
        rowForLoop.forEach((pos, row) -> {
            //Create new row and insert
            XWPFTableRow newRow;
            for (index[0] = dataLength.intValue() - 1; index[0] > 0; index[0]--)
            {
                index[1] = 0;
                newRow = insertTableRow(pos + 1, row, table);
                newRow.getTableCells()
                      .forEach(
                              cell -> {
                                  index[2] = 0;
                                  cell.getParagraphs()
                                      .forEach(
                                              paragraph -> {
                                                  if (paragraph.getRuns()
                                                               .isEmpty())
                                                  {
                                                      return;
                                                  }
                                                  replace.get(index[1])
                                                         .get(index[2]++)
                                                         .forEach(
                                                                 replacement -> replaceParagraph(
                                                                         paragraph,
                                                                         replacement.getStart(),
                                                                         replacement.getEnd(),
                                                                         replacement.valueAt(index[0], dataMap),
                                                                         loopEnd.length()
                                                                 ));
                                              }

                                      );
                                  index[1]++;
                              }
                      );
            }

            //Replace first line
            index[0] = index[1] = 0;
            row.getTableCells()
               .forEach(cell -> {
                   index[2] = 0;
                   cell.getParagraphs()
                       .forEach(
                               paragraph -> replace.get(index[1])
                                                   .get(index[2]++)
                                                   .forEach(
                                                           replacement -> replaceParagraph(
                                                                   paragraph,
                                                                   replacement.getStart(),
                                                                   replacement.getEnd(),
                                                                   replacement.valueAt(index[0], dataMap),
                                                                   loopEnd.length()
                                                           ))

                       );
                   index[1]++;
               });
        });
    }

    /**
     * Copy to generate a new line，and insert row into target position.
     *
     * @param pos   insert position
     * @param row   Source row
     * @param table target table
     * @return new row
     */
    private XWPFTableRow insertTableRow(int pos, XWPFTableRow row, XWPFTable table)
    {
        CTTbl ctTbl = table.getCTTbl();
        int sizeCol = ctTbl.sizeOfTrArray() > 0
                      ? ctTbl.getTrArray(0)
                             .sizeOfTcArray()
                      : 0;
        XWPFTableRow newRow = table.insertNewTableRow(pos);

        //Copy row's style
        newRow.getCtRow()
              .setTrPr(row.getCtRow()
                          .getTrPr());
        XWPFTableCell cell;
        //Copy all cells
        for (int i = 0; i < sizeCol; i++)
        {
            cell = newRow.createCell();
            //Copy cell's style
            cell.getCTTc()
                .setTcPr(row.getCell(i)
                            .getCTTc()
                            .getTcPr());

            //Copy all paragraphs in cell
            for (XWPFParagraph paragraph : row.getCell(i)
                                              .getParagraphs())
            {
                XWPFParagraph newParagraph = cell.addParagraph();
                //Copy paragraph's style
                newParagraph.getCTP()
                            .setPPr(paragraph.getCTP()
                                             .getPPr());

                //Copy all runs in paragraphs
                for (XWPFRun run : paragraph.getRuns())
                {
                    XWPFRun newRun = newParagraph.createRun();
                    //Copy run's style
                    newRun.getCTR()
                          .setRPr(run.getCTR()
                                     .getRPr());
                    //Copy content
                    newRun.setText(run.text());
                }
            }
        }
        return newRow;
    }

    /**
     * Description of the search data loop
     *
     * @param paragraph   Searched paragraph
     * @param list        description list
     * @param dataMap     data map
     * @param methodCache Method for obtaining corresponding fields in cached data
     * @param dataLength  The length of the replaced data
     */
    private void searchLoopVariable(
            XWPFParagraph paragraph,
            List<Replacement> list,
            Map<String, Object> dataMap,
            Map<String, Map<String, Method>> methodCache,
            AtomicInteger dataLength,
            PositionInParagraph startPosition
    )
    {
        //Search variable
        TextSegment start = paragraph.searchText(loopStart, startPosition);
        if (start == null) return;
        TextSegment end = paragraph.searchText(
                loopEnd,
                new PositionInParagraph(start.getEndRun(), start.getEndText(), start.getEndChar())
        );
        if (end == null) return;

        //Find the replaced variable and tag
        String variableKey = paragraph.getText(new TextSegment(
                start.getBeginRun(), end.getEndRun(),
                start.getBeginText(), end.getEndText(),
                start.getBeginChar(), end.getEndChar()
        ));
        variableKey = variableKey.replace(loopStart, "")
                                 .replace(loopEnd, "");

        //Find the replaced variable getting method
        Replacement replacement = new Replacement();
        replacement.setStart(start);
        replacement.setEnd(end);
        replacement.setIndexKey(variableKey);
        if (variableKey.contains("."))
        {
            String[] split = variableKey.split("\\.");
            replacement.setIndexKey(split[0]);
            if (methodCache.containsKey(split[0])
                    && methodCache.get(split[0])
                                  .containsKey(split[1]))
            {
                replacement.setMethod(methodCache.get(split[0])
                                                 .get(split[1]));
            }
            else if (dataMap.containsKey(split[0]))
            {
                Object object = dataMap.get(split[0]);
                if ((object instanceof ArrayList))
                {
                    List<?> curList = (ArrayList<?>) object;
                    if (!curList.isEmpty())
                    {
                        dataLength.set(Math.max(dataLength.intValue(), curList.size()));

                        object = curList.get(0);
                        try
                        {
                            Method method = ClassUtils.methodGetterOf(split[1], object.getClass());
                            replacement.setMethod(method);
                            methodCache.put(split[0], new HashMap<String, Method>()
                            {{
                                put(split[1], method);
                            }});
                        }
                        catch (NoSuchMethodException | NoSuchFieldException e)
                        {
                            logger.error(e.getMessage(), e);
                        }
                    }
                }
            }
        }
        list.add(replacement);

        //Check again if there are any variables in the paragraph to replace
        searchLoopVariable(
                paragraph, list, dataMap, methodCache, dataLength,
                new PositionInParagraph(end.getEndRun(), end.getEndText(), end.getEndChar())
        );
    }

    /**
     * Describe the relevant content of the replacement
     */
    private static class Replacement
    {
        private TextSegment start;
        private TextSegment end;
        private String indexKey;
        private Method method;

        public String valueAt(int index, Map<String, Object> ofData)
        {
            String value = "";
            if (isIndexValue())
            {
                value = String.valueOf(index + 1);
            }
            else if (getMethod() != null)
            {
                try
                {
                    List<?> list = ((List<?>) ofData.get(getIndexKey()));
                    if (list != null && !list.isEmpty() && index < list.size())
                    {
                        value = String.valueOf(
                                getMethod().invoke(list.get(index))
                        );
                    }
                }
                catch (IllegalAccessException | InvocationTargetException e)
                {
                    logger.error(e.getMessage(), e);
                    value = "";
                }
            }
            return value;
        }

        public boolean isIndexValue()
        {
            return "index".equals(indexKey);
        }

        public TextSegment getStart()
        {
            return start;
        }

        public void setStart(TextSegment start)
        {
            this.start = start;
        }

        public TextSegment getEnd()
        {
            return end;
        }

        public void setEnd(TextSegment end)
        {
            this.end = end;
        }

        public String getIndexKey()
        {
            return indexKey;
        }

        public void setIndexKey(String indexKey)
        {
            this.indexKey = indexKey;
        }

        public Method getMethod()
        {
            return method;
        }

        public void setMethod(Method method)
        {
            this.method = method;
        }
    }
}