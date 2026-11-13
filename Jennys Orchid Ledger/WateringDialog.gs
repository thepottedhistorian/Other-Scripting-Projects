<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
        padding: 15px; background-color: #fdfdfd; color: #2c3e50; 
      }
      .container { display: flex; flex-direction: column; gap: 12px; }
      label { font-weight: bold; font-size: 14px; color: #1b5e20; }
      .range-row { display: flex; gap: 8px; }
      input[type="text"], select, textarea { 
        width: 100%; padding: 10px; border: 1px solid #c8e6c9; 
        border-radius: 6px; box-sizing: border-box; font-size: 14px;
      }
      .refresh-btn {
        width: 45px; background-color: #2e7d32; color: white; border: none;
        border-radius: 6px; cursor: pointer; font-size: 16px; 
        display: flex; align-items: center; justify-content: center;
      }
      .refresh-btn:hover { background-color: #1b5e20; }
      textarea { resize: vertical; min-height: 80px; }
      .submit-btn { 
        background-color: #4caf50; color: white; padding: 12px; border: none;
        border-radius: 6px; cursor: pointer; width: 100%; font-weight: bold;
        font-size: 15px; margin-top: 10px;
      }
      .submit-btn:hover { background-color: #388e3c; }
      .hint { font-size: 11px; color: #7f8c8d; margin-top: -8px; }
      .footer-note { font-size: 11px; color: #7f8c8d; text-align: center; margin-top: 10px; }
    </style>
  </head>
  <body>
    <div class="container">
      <div>
        <label for="rangeInput">Apply to Range (Rows)</label>
        <div class="range-row">
          <input type="text" id="rangeInput" placeholder="e.g., 38, 47, 51">
          <button class="refresh-btn" onclick="refreshRange()" title="Refresh selection">🔄</button>
        </div>
      </div>
      <div class="hint">Select rows on sheet, then hit refresh.</div>

      <div>
        <label for="method">Watering Method</label>
        <select id="method">
          <option value="Soak">Soak</option>
          <option value="Mist">Mist</option>
          <option value="Flush">Flush</option>
          <option value="Pour">Pour</option>
          <option value="Spot Watering">Spot Watering</option>
        </select>
      </div>

      <div>
        <label for="notes">Notes</label>
        <textarea id="notes" placeholder="Enter observations..."></textarea>
      </div>

      <button class="submit-btn" id="mainBtn" onclick="submitWatering()">Log Maintenance</button>
      <div class="footer-note">Updates individual specimen ledgers.</div>
    </div>

    <script>
      window.onload = function() { refreshRange(); };

      function refreshRange() {
        google.script.run.withSuccessHandler(function(rangeStr) {
          document.getElementById('rangeInput').value = rangeStr;
        }).getCurrentRangeString();
      }

      function submitWatering() {
        const range = document.getElementById('rangeInput').value;
        const method = document.getElementById('method').value;
        const notes = document.getElementById('notes').value;
        
        if (!range) { alert("Please specify a range."); return; }

        const btn = document.getElementById('mainBtn');
        btn.disabled = true;
        btn.innerText = 'Syncing...';

        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler((err) => {
            alert('Error: ' + err);
            btn.disabled = false;
            btn.innerText = 'Log Maintenance';
          })
          .processBulkWatering({range: range, method: method, notes: notes});
      }
    </script>
  </body>
</html>
