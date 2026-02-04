/**
 * Adds two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

// Essential: This initializes the functions
CustomFunctions.associate("ADD", add);