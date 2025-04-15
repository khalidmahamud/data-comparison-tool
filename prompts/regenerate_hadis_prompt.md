<!--
This script uses these variables:
1. {hadis_arabic_text} - The original Arabic Hadith text
2. {hadis_translated_bangla} - The Bengali translation that may need correction
3. {current_analysis} - Any previous analysis if available
-->

## Arabic Hadith:

{{{hadis_arabic_text}}}

## Bengali Translation:

{{{hadis_translated_bangla}}}

## Instructions:

1. Compare the Arabic and Bengali texts:

   - If they don't match semantically, respond ONLY with: "Arabic and Bangla don't match"
   - Otherwise, proceed with corrections

2. If they match, fix these issues while preserving the exact meaning:

   - Correct outdated Bengali spellings (ওয়াহী → ওহী, ফিরিশতা → ফেরেশতা, etc.)
   - Fix conjunct letter errors (কাফ্ফারা → কাফ্‌ফারা, আম্র → আম্‌র, etc.)
   - Correct companion names if misspelled
   - Remove unnecessary special characters from words and names
   - Fix any missing text or incorrect translations that don't match the Arabic
   - Fix inconsistent use of special characters: Standardize all instances where apostrophes ('), hyphens (-), and special character combinations disrupt readability while preserving the original meaning of the text.

3. DO NOT:

   - Do Not Change honorifics like (সাল্লাল্লাহু আলাইহি ওয়া সাল্লাম), (রাঃ), etc.
   - Do Not Modify any text marked with (আরবি) or (আরবী)
   - Do Not Add any Arabic text that wasn't in the original translation
   - Do Not Add any explanatory text or comments
   - Do Not Add special characters unless they existed in the original or are grammatically required

## Output:

Return ONLY the corrected Bengali translation without any additional text.
