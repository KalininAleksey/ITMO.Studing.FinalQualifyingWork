$(document).ready(function(){


  $(document).on("contextmenu",function(){
    return false;
  });

  $(".menu-main").on("click",function(e){
    var currentitem=$("body").find(".current");
    var newcurrentclass=$(e.target).attr('class');
    var newcurrentitem=$("body").find("."+newcurrentclass);
    if ((newcurrentclass!="profile-menu") && (newcurrentclass!="menu-main")) {
       for (i=0;i<currentitem.length;i++){
        $(currentitem[i]).removeClass("current");
        }
    }
    if (newcurrentclass!="menu-main")
    {
        for (i=0;i<newcurrentitem.length;i++){
        $(newcurrentitem[i]).addClass("current");
        }
     }
  });

  // Выпадающее меню
  /*$(".profile-menu").on("click",function(){
    var submenu = $(this).parent().find(".submenu");
    submenu.addClass("active");
    submenu.fadeIn(300);

    $(document).on("mousedown",function(e){
      if($(e.target).attr('class')!="profile-menu")
      {
      submenu.removeClass("active");
      submenu.fadeOut(300);
      }
    })
  });*/
  $("#fileuploader").on("change",function(){
     var file = this.files[0];
        if((file.size > 20848820) && ($(this).val().split('.').pop().toLowerCase() != "docx")){
                  $("#sendformbtn").attr('disabled', true);
                  alert("Выбирать можно только файлы с расширением docx и размером не более 20 МБ");
        }

       if(file.size > 20848820){
          $("#sendformbtn").attr('disabled', true);
          alert("Размер файла должен быть не более 20 МБ");
       }
        else if ($(this).val().split('.').pop().toLowerCase() != "docx") {
          $("#sendformbtn").attr('disabled', true);
          alert("Выбирать можно только файлы с расширением docx");
       }
       else{
        $("#sendformbtn").attr('disabled', false);
       }
    });

})