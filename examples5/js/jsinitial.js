//jQuery.noConflict();
(function($) { 
if ( $(window).width() > 1024) {
jQuery(document).ready(function($) {
	$('html').localScroll(800);

	//.parallax(xPosition, speedFactor, outerHeight) options:
	//xPosition - Horizontal position of the element
	//inertia - speed to move relative to vertical scroll. Example: 0.1 is one tenth the speed of scrolling, 2 is twice the speed of scrolling
	//outerHeight (true/false) - Whether or not jQuery should use it's outerHeight option to determine when a section is in the viewport
	$('#banner-wrap').parallax("50%", 0.5);
	$('#feedback-wrap').parallax("50%", 0.5);
	
	$(window).scroll(function() {
		if($(this).scrollTop() != 0) {
			$('#toTop').fadeIn();
		} else {
			$('#toTop').fadeOut();
		}
	});
 
	$('#toTop').click(function() {
		$('body,html').animate({scrollTop:0},800);
	});
			
})

new WOW().init();

// Set options
var options = {
	offset: '#showHere',
	classes: {
		clone:   'banner--clone',
		stick:   'banner--stick',
		unstick: 'banner--unstick'
	}
};

// Initialise with options
var banner = new Headhesive('#nav-wrap', options);

// Headhesive destroy
// banner.destroy();
}
else {
jQuery(document).ready(function($) {
	$('html').localScroll(800);

	//.parallax(xPosition, speedFactor, outerHeight) options:
	//xPosition - Horizontal position of the element
	//inertia - speed to move relative to vertical scroll. Example: 0.1 is one tenth the speed of scrolling, 2 is twice the speed of scrolling
	//outerHeight (true/false) - Whether or not jQuery should use it's outerHeight option to determine when a section is in the viewport
	
	$(window).scroll(function() {
		if($(this).scrollTop() != 0) {
			$('#toTop').fadeIn();
		} else {
			$('#toTop').fadeOut();
		}
	});
 
	$('#toTop').click(function() {
		$('body,html').animate({scrollTop:0},800);
	});
			
})

new WOW().init();

// Set options
var options = {
	offset: '#showHere',
	classes: {
		clone:   'banner--clone',
		stick:   'banner--stick',
		unstick: 'banner--unstick'
	}
};

// Initialise with options
var banner = new Headhesive('#nav-wrap', options);

// Headhesive destroy
// banner.destroy();
}
})(jQuery);