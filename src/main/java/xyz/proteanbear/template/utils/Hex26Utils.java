package xyz.proteanbear.template.utils;

/**
 * Hex26(A,B,C,……,Z,…… like Excel's column) exchange Hex 10 tools
 *
 * @author ProteanBear
 */
public class Hex26Utils
{
    //All numbers
    private static final String[] numbers=new String[]{"","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};

    /**
     * Hex 10 to Hex 26<br/>
     * Number must be between 1 and 676
     *
     * @param hex10 the number of hex 10
     * @return the number string of hex 26,eg.A,B,C
     */
    public static String from(int hex10)
    {
        if(hex10<1) return hex10+"";
        if(hex10>26*26) return hex10+"";

        StringBuilder result=new StringBuilder();
        int multiple=Math.max((hex10-1)/26,0), remainder=hex10%26;
        result.append((multiple==0)?"":numbers[multiple])
                .append(remainder==0?numbers[26]:numbers[remainder]);

        return result.toString();
    }
}