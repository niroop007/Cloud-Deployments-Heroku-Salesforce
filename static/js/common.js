function hide_Func() {
		  var checkBox1 = document.getElementById("B1");
		  var checkBox2 = document.getElementById("B2");
		  var checkBox3 = document.getElementById("B3");
		  var staircase1 = document.getElementById("SC1");
		  var input1 = document.getElementById("b1_input");
		  var wattcheck1=document.getElementById("b1_watt")
		  var input2 = document.getElementById("b2_input");
		  var wattcheck2=document.getElementById("b2_watt")
		  var input3 = document.getElementById("b3_input");
		  var wattcheck3=document.getElementById("b3_watt")
		  var input4 = document.getElementById("sc1_input");
		  var stairwattcheck1=document.getElementById("sc1_watt")
		  if (checkBox1.checked == true){
			input1.style.display = "block";
			wattcheck1.style.display = "block";
		  }
			else if(checkBox2.checked == true){
			input2.style.display = "block";
			wattcheck2.style.display = "block";
			}
			else if(checkBox3.checked == true){
			input3.style.display = "block";
			wattcheck3.style.display = "block";
			}
			else if(staircase1.checked == true){
			input4.style.display = "block";
			stairwattcheck1.style.display = "block";
			}
		   else {
			input1.style.display = "none";
			wattcheck1.style.display = "none";
			input2.style.display = "none";
			wattcheck2.style.display = "none";
			input3.style.display = "none";
			wattcheck3.style.display = "none";
			input4.style.display = "none";
			stairwattcheck1.style.display = "none";
		  }
		}
		function newFunction() {
            document.getElementById("quote").reset();
         }