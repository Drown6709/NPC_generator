# NPC_generator_excel
Just another random NPC generator - This time in excel.

This is just a simple excel file for generating random NPCs. It is not refined, beautiful or particularly well-coded. In fact it's the opposite. But it was done quickly and it works for my own personal usage with all the imperfections it has.

**Fair Warning: Do not think this will work for you.** It might, but the intended audience is just me. So I'm only sharing it because I figured others might find it useful. It comes with zero guarantees of future updates, bug correction or any improvements.

The basic is a few sheets displaying a random NPC; both completely random and a few specific classes or races. I might add more if I find the need. I might not.

**Important functionality instruction 1: Calculation should be set to "Manual"** the workbook is meant to be used in Excel with the auto-calculation off! This can be set to on or off under options and calculation options. If it's set to automatic, the NPCs will re-roll every time you do anything in the workbook. Just be aware that this is a global setting in Excel, meaning that ALL your excel files will calculate manually while this is set.

**Important functionality instruction 2:The workbook uses a macro** Since auto-calculation is meant to be turned off, there's a button that when clicked re-rolls the NPC, but only the active NPC. Not the other sheets. However, this is a macro, so for it to work you need to enable macros. If you do not feel comfortable doing that, you can use the "Calculate Sheet" function under the "Formula" ribbon. However, this will not take into account the plot hook. The plot hook is based on a calculation on the data sheet that is also triggered by re-rolling an NPC. So you'll have to either calculate the entire workbook, or calculate both the NPC sheet and the data sheet to re-roll a plot hook.

**DnD setting.** The data is meant to be used in the Mystara setting, specifically Karameikos. This mostly refers to names and races. Feel free to just switch out those in the tables to better fit your setting.

**So how is it intended to be used?** The idea is that in the sheet called "Random NPC Source Data" is a ton of tables with various values. On all the other sheets are a series of cells which is populated with random data from the tables correspoding to the type of data. I think it's fairly self-explanatory once you open the workbook. So when you click the button "Create Random NPC" all the cells will re-roll and pull new random data (Provided you've activated macros).

You can use it as it is, and never change anything. If you wish to save an NPC, you can either copy the data over into a new workbook (Remember to paste only values and not a regular paste!), or ou can simply create a copy of the worksheet and store it within the workbook, renaming the sheet name to something meaningful. Just remember that you need to set calculation to manual, or it will re-roll your saved NPC every time you do anything.

It is not meant to be a complete NPC creator. Just something I use in a pinch when I need one.

If you wish to change the data for a category, say the hair, you go into "Random NPC Source Data" and look up the corresponding column in the corresponding table and change to your heart's delight. You can add, delete or change content all you want as long as you don't break the table itself, you should be fine.

For plot hooks it works a little different. A plot hook has both a hook, and a category. The reason is that this way you can adjust the likelyhood of a specific hook being rolled. So first it picks a random category, and then it picks a random row within that category. This means that if you have a category with tons of rows and a category with few rows, the probability of picking a specific hook from the large category goes down. Simply because there're more to pick from in the large category than the small. So each category has an equal chance of getting picked. But overall it's more likely to pick a specific line from a small category than a large category. I do this because some of the hooks are very much the same with just a slight variation. Like "NPC is missing a precious coin" and "NPC is missing a precious necklace" etc. So I've just created a category that lumps all "NPC is missing a precious XXX" into one, and then created a different category with "NPC is looking for a long lost son", "NPS is looking for a long lost parent" etc. So it's equally likely to pick either missing something or looking for someone, even though there's a lot more missing something-hooks. I hope this makes sense.

**Content source:** The content of the tables have been pulled from various open sources around the internet. Here's a list of links to the ones I remember. Should you find your content here, and your credit not mentioned, DM me and I'll make sure to give you the right credit.

**Content sources:**

https://github.com/cellule/dndGenerator/

https://www.reddit.com/user/OrkishBlade/

https://github.com/Tetra-cube/Tetra-cube.github.io

plus a few flavours of my own, probably brorrowed/inspired by soneone else I can't remember...
