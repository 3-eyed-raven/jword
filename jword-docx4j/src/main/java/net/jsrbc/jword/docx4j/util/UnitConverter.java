package net.jsrbc.jword.docx4j.util;

/**
 * 单位转换
 * @author ZZZ on 2021/2/2 16:52
 * @version 1.0
 */
public final class UnitConverter {
    /**
     * 厘米转twips
     * @param cmValue 厘米值
     * @return twips值
     */
    public static long cmToTwips(double cmValue) {
        return Math.round(cmValue * 567);
    }

    /**
     * 百分比转百分之五十值
     * @return 百分比
     */
    public static int percentToFifthPct(double percentValue) {
        return (int) Math.round(percentValue * 50);
    }

    /**
     * 厘米转English Metric Units
     * @param cmValue 厘米值
     * @return English Metric Units值
     */
    public static long cmToEMUs(double cmValue) {
        return Math.round(cmValue * 360000);
    }

    private UnitConverter() {}
}
