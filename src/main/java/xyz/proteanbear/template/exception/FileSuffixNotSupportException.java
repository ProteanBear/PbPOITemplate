package xyz.proteanbear.template.exception;

/**
 * Throw a exception when file suffix is not supported.
 */
public class FileSuffixNotSupportException extends Exception
{
    public FileSuffixNotSupportException()
    {
        super("The file suffix is not supported");
    }
}