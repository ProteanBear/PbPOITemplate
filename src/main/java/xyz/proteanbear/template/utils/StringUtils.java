package xyz.proteanbear.template.utils;

/**
 * String tool
 *
 * @author ProteanBear
 */
public class StringUtils
{
    /**
     * Check if the string is blank.
     *
     * @param string the string.
     * @return If the string is blank,return true.
     */
    public static boolean isBlank(String string)
    {
        return (string == null || "".equals(string.trim()));
    }
}