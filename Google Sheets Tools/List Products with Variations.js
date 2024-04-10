function count_variations_with_ids(data) {
  
    // This function takes a list of lists containing Shopify product and variant data and returns the number of products with variations and a list of those product IDs.
    
    // Args:
          // data: A list of lists containing product and variant data.
    
      // Returns:
        // A tuple containing:
          //   - The number of products with variations.
          //   - A list of product IDs that have variations.


    // In your spreadsheet, you can call the function like this (assuming the data is in sheet1 and starts from row 2): =count_variations_with_ids(sheet1!A2:B)
    // This function to count the number of products with variations and create a list of those product IDs:

    
      product_counts = {};
      variation_products = [];
      for (const row of data) {
        const product_id = row[0];
        if (product_id in product_counts) {
          product_counts[product_id] += 1;
        } else {
          product_counts[product_id] = 1;
        }
      }
      variation_count = 0;
      for (const product_id of Object.keys(product_counts)) {
        if (product_counts[product_id] > 1) {
          variation_count += 1;
          variation_products.push(product_id);
        }
      }
      return [variation_count, variation_products];
    }
    