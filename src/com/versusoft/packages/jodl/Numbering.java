/**
 *  odt2daisy - OpenDocument to DAISY XML/Audio
 *
 *  (c) Copyright 2008 - 2009 by Vincent Spiewak, All Rights Reserved.
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
public class Numbering {

    private static final String[] LETTER_NUMS = {"", "a", "b", "c", "d", "e", "f", "g", "h", "i",
        "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
    };
    private static final String[] ROMAN_NUMS = {"m", "cm", "d", "cd", "c", "xc", "l",
        "xl", "x", "ix", "v", "iv", "i"
    };
    private static final int[] NUMRERAL_NUMS = {1000, 900, 500, 400, 100, 90, 50,
        40, 10, 9, 5, 4, 1
    };

    /**
     * ex: A, B, ..., Z, AA, AB, ..., AZ, BA, BB, BC ...
     * @param number
     * @return
     */
    public static String toLetter(int number) {

        
        String letter = "";
        
         if(number < 27){
            letter += LETTER_NUMS[number];
        } else if(number < 709){
        int div = number / 26;
         letter = LETTER_NUMS[div]+LETTER_NUMS[number-(26*div)];
        } 
     /*   
        if(number < 27){
            letter += LETTER_NUMS[number];
        } else if(number < 53){
         letter = 'a'+LETTER_NUMS[number-26];
        } else if(number < 79){
            letter = 'b'+LETTER_NUMS[number-52];
        } else if(number < 105){
            letter = 'c'+LETTER_NUMS[number-78];
        } else if(number < 131){
            letter = 'd'+LETTER_NUMS[number-104];
        } else if(number < 157){
            letter = 'e'+LETTER_NUMS[number-130];
        } else if(number < 183){
            letter = 'f'+LETTER_NUMS[number-156];
        } else if(number < 209){
            letter = 'g'+LETTER_NUMS[number-182];
        } else if(number < 235){
            letter = 'h'+LETTER_NUMS[number-208];
        } else if(number < 261){
            letter = 'i'+LETTER_NUMS[number-234];
        } else if(number < 287){
            letter = 'j'+LETTER_NUMS[number-260];
        } else if(number < 313){
            letter = 'k'+LETTER_NUMS[number-286];
        } else if(number < 339){
            letter = 'l'+LETTER_NUMS[number-312];
        } else if(number < 365){
            letter = 'm'+LETTER_NUMS[number-338];
        } else if(number < 391){
            letter = 'n'+LETTER_NUMS[number-364];
        } else if(number < 417){
            letter = 'o'+LETTER_NUMS[number-390];
        } 
*/
        /*
         else if(number < 702 ){
            int i = number / 26;
            letter = LETTER_NUMS[i] + LETTER_NUMS[number - (i*26)];
        }
         */
        /*while (number > 0) {

            if (number < 27) {
                letter += LETTER_NUMS[number];
                number = 0;
                
            } else {
                
                if(letter.charAt(cursor) == 'z'){
                    
                }
                    
                    letter += ref;
                    number -= 26;
            }
        }*/

        return letter;
    }

    public static String toRoman(int number) {
        if (number <= 0 || number >= 4000) {
            throw new NumberFormatException("Value outside roman numeral range.");
        }
        String roman = "";         // Roman notation will be accumualated here.

        for (int i = 0; i < ROMAN_NUMS.length; i++) {
            while (number >= NUMRERAL_NUMS[i]) {
                number -= NUMRERAL_NUMS[i];
                roman += ROMAN_NUMS[i];
            }
        }
        return roman;
    }
}
