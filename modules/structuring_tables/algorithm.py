# Input: A list of sorted natural numbers in ascending order, containing several consecutive series of numbers
    # Example: [1,2,3,10,11,20,21,22,23,24,25]
# Output: A list of strings where each element is a concise expression of the corresponding consecutive series
    # Example: ['1:3','10:11','20:25']
    # Note: Different from definition of slices in Python, here the right end number is included.
def get_consecutive_ranges(list):
    list.sort(reverse=False) # Sort the list again just in case
    output = [] # Initiate output list (of strings)
    start = 0 # By default the slice starts from the very first item

    for i in range(1, len(list)):
        if list[i] != list[i-1] + 1: # Identify the cutpoints
            current_slice = f"{list[start]}:{list[i-1]}" # Construct the current slice
            output.append(current_slice) # Append the current slice to output list
            start = i # Start of the next slice
    current_slice = f"{list[start]}:{list[len(list)-1]}" # Not to forget to construct the final slice
    output.append(current_slice)
    
    return output

