/* Sample workbook: two sheets (Cars, Specs), English, small. */
window.SAMPLE_SHEETS = [
  {
    name: "Cars",
    rows: [
      ["Brand", "Model"],
      ["Audi", null],
      ["Audi", "A3"],
      ["Audi", "A4"],
      ["Audi", "Q5"],
      ["BMW", null],
      ["BMW", "3 Series"],
      ["BMW", "5 Series"],
      ["BMW", "X5"],
      ["Ford", null],
      ["Ford", "Focus"],
      ["Ford", "Mustang"],
      ["Honda", null],
      ["Honda", "Civic"],
      ["Honda", "CR-V"],
      ["Tesla", null],
      ["Tesla", "Model 3"],
      ["Tesla", "Model Y"],
      ["Toyota", null],
      ["Toyota", "Camry"],
      ["Toyota", "Corolla"],
      ["Toyota", "RAV4"]
    ]
  },
  {
    name: "Specs",
    rows: [
      ["Brand", "Country", "Founded", "Avg. Price (USD)"],
      ["Audi", "Germany", 1909, 45000],
      ["BMW", "Germany", 1916, 52000],
      ["Ford", "USA", 1903, 38000],
      ["Honda", "Japan", 1948, 30000],
      ["Tesla", "USA", 2003, 48000],
      ["Toyota", "Japan", 1937, 31000]
    ]
  }
];

/* Backward compatibility: first sheet only, same shape as before. */
window.SAMPLE_SHEET = window.SAMPLE_SHEETS[0];
