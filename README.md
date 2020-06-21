# MonsterSudokuNov This was a project for a senior, me, to exercise his brain by working regulalar Sudoku 
# puzzles. But, I soon found that it takes far too long to write in the numbers that were still possible 
# -- the remaining numbers. So, instead of spending a lot of time writing in numbers, I decided to create 
# a program that would do it for me. After that worked, I found that there were some cells that had only 
# one number remaining. So, I wrote a procedure that found those numbers and completed them. Since then,  
# I have found other patterns that allow the program to input answers or to simplify the puzzle by reducing 
# the number remaining numbers. Along the way, I found the much larger Monster Sudoku puzzles; so I created
# /modified my program to work them. Now, after you input the starting numbers, the program can complete 
# some 5 star Monster Sudoku puzzles just by activating some of the button procedures. The point is that 
# I didn't know how to do some of these things until after I started writing the program. I hope others 
# can read the code, I am not a programmer, and see how things work. There is no guessing or matrix 
# manipulation. Every routing is designed to find a specific pattern of numbers, and to use those numbers
# to either find a solution to a cell, or to simplify the puzzle. I have improved some of the original 
# prodecures to make them run a littele faster, but I didn't spend time just trying to optimize the
# solution. The program is a memory hog!
# *** Contributions are welcome. In order to save the puzzle, I save array data to a textbox, cut the text,
# and paste it into the code in a specific "solution" procedure. Terrible! It is easy to save that solution
# to a file, but I haven't figured out how to read that file so I can recreate the solution by updating the
# working array. It would help if someone would do that. I don't want to require a database connection to
# a database -- which a different user might not have. 
# If you use or review the program, please provide feedback on your thoughts. Remember, I am not a programmer;
# I am a hobby programmer. 
