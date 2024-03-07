var grossWeightAutosum=0;
var totalVolume = 0;
document.getElementById('submitButton').addEventListener('click', function() {
    var fileInput = document.getElementById('fileInput');
    if (fileInput.files.length > 0) {
        var file = fileInput.files[0];
        var reader = new FileReader();
        
        reader.onload = function(event) {
            var data = new Uint8Array(event.target.result);
            var workbook = XLSX.read(data, {type: 'array'});
            
            // Assuming the first sheet contains the dimensions
            var sheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[sheetName];
            
            // Assuming dimensions are stored in column 'E' (5th column) starting from the 2nd row
            var rowNum = 2; // Starting row number
            var dimensions = [];
            
            while (worksheet['E' + rowNum]) {
                var dimensionCell = worksheet['E' + rowNum].v; // Assuming data is stored directly in the cell
                var dimensionParts = dimensionCell.split('*'); // Split dimension and multiplier
                var dimension = dimensionParts[0].trim();
                var multiplier = dimensionParts[1] ? parseInt(dimensionParts[1].trim()) : 1; // If no multiplier specified, default to 1
                
                // Repeat the dimension based on the multiplier
                for (var i = 0; i < multiplier; i++) {
                    dimensions.push(dimension);
                }
                
                rowNum++;
            }
            // Save dimensions in a variable
            window.dimensionsData = dimensions;
            
            // Calculate and display total volume
            calculateAndDisplayTotalVolume(dimensions);
            grossWeightAutosum = calculateGrossWeightAutosum(worksheet);
           

            var weight_show = document.getElementById('weight');
            weight_show.innerHTML = '';
            var total_Weight_TextNode = document.createTextNode(`Autosum of Gross Weight: ${grossWeightAutosum}`);
            weight_show.appendChild(total_Weight_TextNode);
        };
        
        reader.readAsArrayBuffer(file);
    } else {
        alert('Please select a file.');
    }
});

// Function to calculate total volume of all boxes and display it
function calculateAndDisplayTotalVolume(dimensions) {
    
  
    
    for (var i = 0; i < dimensions.length; i++) {
        var dimension = dimensions[i];
        var dimensionWithoutCM = dimension.replace(/CM/g, '').trim();
        var dimensionParts = dimensionWithoutCM.split('X').map(function(part) {
            return parseInt(part.trim());
        });
        
        if (dimensionParts.length === 3) {
            var length = dimensionParts[0];
            var width = dimensionParts[1];
            var height = dimensionParts[2];
            var volume = length * width * height;
            totalVolume += volume;
        }
    }
    var Volume_show = document.getElementById('Volume');
    Volume_show.innerHTML = '';
    var total_Volume_TextNode = document.createTextNode(`Total: ${totalVolume}`);
    Volume_show.appendChild(total_Volume_TextNode);
    console.log("hello")
}

// Function to calculate autosum of Gross Weight column
function calculateGrossWeightAutosum(worksheet) {
    var autosum = 0;
    var columnIndex = 1; // Assuming Gross Weight is in the 2nd column (0-based index)

    // Parse sheet data into array
    var sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Iterate over each row (starting from the 2nd row)
    for (var i = 1; i < sheetData.length; i++) {
        var row = sheetData[i];
        
        // Check if the row has data for the Gross Weight column
        if (row && row[columnIndex] != null) {
            // Parse the value and add it to the autosum
            var weight = parseFloat(row[columnIndex]);
            if (!isNaN(weight)) {
                autosum += weight;
            }
        }
    }

    return autosum;
}


document.getElementById('FindVehicle').addEventListener('click', function() {
    fetch('vehicle detail.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            var workbook = XLSX.read(data, { type: 'arrayBuffer' });
            var len=0;
            // Assuming the first sheet contains the data
            var sheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[sheetName];

            // Extract data from the C and D columns
            var FColumnData = [];
            var GColumnData = [];

            for (var i = 1; ; i++) {
                var FCellValue = worksheet['F' + i];
                var GCellValue = worksheet['G' + i];
                len++;
                if (!FCellValue && !GCellValue) {
                    // No more data in both columns, exit the loop
                    break;
                }

                // Extract the cell values
                var cValue = FCellValue ? FCellValue.v : '';
                var dValue = GCellValue ? GCellValue.v : '';

                // Push the values to the arrays
                FColumnData.push(cValue);
                GColumnData.push(dValue);
            }

            // Display the extracted data in the console
            // console.log('Data from C column:', cColumnData);
            // console.log('Data from D column:', dColumnData);
            for(var i=1;i<=len;i++){
                if(GColumnData[i]>totalVolume && FColumnData[i]>grossWeightAutosum){
                    console.log("row",GColumnData[i],i+1);
                    
                        console.log("final row: ", i+1);
                        // const valumneValue = worksheet['G' + i+1];
                        var j=i+1
                        var cellValue = worksheet['B' + j]; // Construct cell reference dynamically
                        // const vahicaleName= worksheet['B' + i+1];
                        // console.log('Value in cell G' + i+1 + ':', valumneValue);
                        // console.log('Value in cell F' + i+1+ ':', cellValue);
                        var Vehicle = document.getElementById('Vehicle');
                        var wTextNode = document.createTextNode( cellValue.w);

                        Vehicle.appendChild(wTextNode);
                        console.log('car: ',cellValue);
                        break;
                }
                
                // console.log("Less Than ",i);
            }
        })
        .catch(error => {
            console.error('Error fetching Excel data:', error.stack);
        });
});
