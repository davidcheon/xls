$(function(){
$("#submit").click(
	function (evt){
		var form=document.form;
		var startid=parseFloat(form.startid.value);
		var endid=parseFloat(form.endid.value);
		var startphonenum=parseFloat(form.startphonenum.value);
		var savefilepath=form.savefilepath.value.trim();
		var savecountsperfile=parseFloat(form.savecountsperfile.value);
		if(isNaN(startid) ||isNaN(endid)||isNaN(startphonenum)||isNaN(savecountsperfile)|| savefilepath=="" ){

			alert("Invalid format for startid or endid or startphonenum or savecountsperfile");
		}
		else if(startid>endid){
		alert("StartID can not be greater than EndID");
		}
		else{
		obj=$.ajax({
		type:"post",
		async:false,
		url:"xls.py/handler",
		timeout:3000,
		data:$("#form").serialize(),
		cache:false,
		dataType:"json",
		success:function(data,statuss){
				if (statuss=="success"){
				alert(data.info);
				$("#statusbar").empty();
				$("#statusbar").append(data.info);
				$("#statusbar").css("color","red");
				}
			},
		});
	    }

	}
	)
}
);
