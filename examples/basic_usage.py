"""Basic usage example for html2docx."""

from html2docx import HTMLToDocx

html = """
<!DOCTYPE html>
<html>
<head>
<style>
:root {
  --heading-color: #2c3e50;
  --accent: #e74c3c;
}
</style>
</head>
<body>
  <h1 style="color: var(--heading-color)">Quarterly Report</h1>
  <p>This report covers <strong>Q1 2026</strong> performance metrics.</p>

  <h2>Summary</h2>
  <p>Revenue increased by <em>15%</em> compared to the previous quarter.</p>

  <h2>Key Metrics</h2>
  <table>
    <thead>
      <tr><th>Metric</th><th>Q4 2025</th><th>Q1 2026</th></tr>
    </thead>
    <tbody>
      <tr><td>Revenue</td><td>$50,000</td><td>$57,500</td></tr>
      <tr><td>Users</td><td>1,200</td><td>1,450</td></tr>
      <tr><td>Churn</td><td>5.2%</td><td>4.1%</td></tr>
    </tbody>
  </table>

  <h2>Action Items</h2>
  <ol>
    <li>Review pricing strategy</li>
    <li>Launch email campaign</li>
    <li>Update onboarding flow</li>
  </ol>

  <h2>Notes</h2>
  <blockquote>
    <p>Focus on retention over acquisition this quarter.</p>
  </blockquote>

  <pre><code>SELECT metric, value FROM quarterly_report WHERE quarter = 'Q1-2026';</code></pre>
</body>
</html>
"""

converter = HTMLToDocx()
output = converter.convert(html, "example_report.docx")
print(f"Converted to: {output}")
print(f"Stats: {converter.stats}")
