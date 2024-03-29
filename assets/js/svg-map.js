$(document).ready(function(){
    //Initialize bootstrap popovers
    $('[data-toggle="popover"]').popover();

    // Initialize bootstrap carousel
    $('.carousel').carousel({interval: false})

    // Check if it is iOS
    var isiOS = (navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false);
    
    if(isiOS === true) {
        // Store -webkit-tap-highlight-color as this gets set to rgba(0, 0, 0, 0) in the next part of the code
        var tempCSS = $('a').css('-webkit-tap-highlight-color');
        
        $('body').css('cursor', 'pointer')                                    // Make iOS honour the click event on body
                 .css('-webkit-tap-highlight-color', 'rgba(0, 0, 0, 0)');     // Stops content flashing when body is clicked
             
        // Re-apply cached CSS
        $('a').css('-webkit-tap-highlight-color', tempCSS);
    }

    $('body').on('click', function (e) {
        $('[data-toggle="popover"]').each(function () {
            //the 'is' for buttons that trigger popups
            //the 'has' for icons within a button that triggers a popup
            if (!$(this).is(e.target) && $(this).has(e.target).length === 0 && $('.popover').has(e.target).length === 0) {
                $(this).popover('hide');
            }
        });

        $('body').on('hidden.bs.popover', function() {
            var popovers = $('.popover').not('.in');
            if (popovers) {
                popovers.remove();
            }
        });
    });
});

//remove popovers - used for when window resizes
function removePopovers() {
    var popovers = $('.popover');
    if (popovers) {
        popovers.remove();
    }
}