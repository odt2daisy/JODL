/**
 *  odt2daisy - OpenDocument to DAISY XML/Audio
 *
 *  (c) Copyright 2008 - 2012 by Vincent Spiewak, All Rights Reserved.
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Lesser Public License as published by
 *  the Free Software Foundation; either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public License along
 *  with this program; if not, write to the Free Software Foundation, Inc.,
 *  51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
 */
package com.versusoft.packages.jodl;

/**
 *
 * @author Vincent Spiewak
 */
public class LetterNumbering {

    private static final String[] LETTER_NUMS = { "a", "b", "c", "d", "e", "f", "g", "h", "i",
        "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
    };


    /**
     * Convert number to letter in base 26
     *
     * ex: A, B, ..., Z,
     *     AA, AB, ..., AZ,
     *     BA, BB, BC ... BZ,
     *     ...,
     *     ZA, ZB, ZC ... ZZ,
     *     AAA, AAB,..., AAZ,
     *     ABA, ABB, ABC, ... ABZ,
     *     ...
     *
     * @param number (MUST be > 1)
     * @return letter
     *
     */
    public static String toLetter(int n){
        if(n<1) 
            return null;
        else
            return toLetterSub(n-1);
    }

    private static String toLetterSub(int n) {

        int r = n % 26;
        String result = "";

        if (n - r == 0) {
            result = LETTER_NUMS[n];
        } else {
            result = toLetterSub(((n - r) - 1) / 26) + LETTER_NUMS[r];
        }

        return result;
    }

    /** Main method for testing purpose */
    public static void main(String args[]) {
        int limit = 1320;
        for (int i = 1; i < limit; i++) {
            System.out.println(i + ": " + toLetter(i));
        }
    }
}
