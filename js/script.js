jQuery(function ($) { "use strict";
	
	/* ========================================================================= */
	/*	Page Preloader
	/* ========================================================================= */
	
	window.onload = function () {
		document.getElementById('preloader').style.display = 'none';
	}


	/* ========================================================================= */
	/*	Post image slider
	/* ========================================================================= */
	
	$("#post-thumb, #gallery-post").slick({
		infinite: true,
		arrows:false,
		autoplay: true,
  		autoplaySpeed: 4000
		
	});
	
	$("#features").slick({
		infinite: true,
		arrows:false,
		autoplay: true,
  		autoplaySpeed: 4000
	});


	/* ========================================================================= */
	/*	Menu item highlighting
	/* ========================================================================= */


	$("#navigation").sticky({
		topSpacing : 0
	});
	

	/* ========================================================================= */
	/*	Magnific popup
	/* =========================================================================  */
	$('.image-popup').magnificPopup({
    type: 'image',
    removalDelay: 160, //delay removal by X to allow out-animation
    callbacks: {
        beforeOpen: function () {
            // just a hack that adds mfp-anim class to markup
            this.st.image.markup = this.st.image.markup.replace('mfp-figure', 'mfp-figure mfp-with-anim');
            this.st.mainClass = this.st.el.attr('data-effect');
        }
    },
    closeOnContentClick: true,
    midClick: true,
    fixedContentPos: false,
    fixedBgPos: true
});
	/* ========================================================================= */
	/*	Portfolio Filtering Hook
	/* =========================================================================  */

  	var mixer = mixitup('.portfolio-items-wrapper');
	
	/* ========================================================================= */
	/*	Testimonial Carousel
	/* =========================================================================  */
 
	//Init the carousel
	$("#testimonials").slick({
		infinite: true,
		arrows:false,
		autoplay: true,
  		autoplaySpeed: 4000
	});






	/* ========================================================================= */
	/*   Contact Form Validating
	/* ========================================================================= */


	$('#contact-submit').click(function (e) {

		//stop the form from being submitted
		e.preventDefault();

		/* declare the variables, var error is the variable that we use on the end
		to determine if there was an error or not */
		var error = false;
		var name = $('#name').val();
		var email = $('#email').val();
		var subject = $('#subject').val();
		var message = $('#message').val();

		/* in the next section we do the checking by using VARIABLE.length
		where VARIABLE is the variable we are checking (like name, email),
		length is a JavaScript function to get the number of characters.
		And as you can see if the num of characters is 0 we set the error
		variable to true and show the name_error div with the fadeIn effect. 
		if it's not 0 then we fadeOut the div( that's if the div is shown and
		the error is fixed it fadesOut. 
		
		The only difference from these checks is the email checking, we have
		email.indexOf('@') which checks if there is @ in the email input field.
		This JavaScript function will return -1 if no occurrence have been found.*/
		if (name.length == 0) {
			var error = true;
			$('#name').css("border-color", "#D8000C");
		} else {
			$('#name').css("border-color", "#666");
		}
		if (email.length == 0 || email.indexOf('@') == '-1') {
			var error = true;
			$('#email').css("border-color", "#D8000C");
		} else {
			$('#email').css("border-color", "#666");
		}
		if (subject.length == 0) {
			var error = true;
			$('#subject').css("border-color", "#D8000C");
		} else {
			$('#subject').css("border-color", "#666");
		}
		if (message.length == 0) {
			var error = true;
			$('#message').css("border-color", "#D8000C");
		} else {
			$('#message').css("border-color", "#666");
		}

		//now when the validation is done we check if the error variable is false (no errors)
		if (error == false) {
			//disable the submit button to avoid spamming
			//and change the button text to Sending...
			$('#contact-submit').attr({
				'disabled': 'false',
				'value': 'Sending...'
			});

			/* using the jquery's post(ajax) function and a lifesaver
			function serialize() which gets all the data from the form
			we submit it to send_email.php */
			$.post("sendmail.php", $("#contact-form").serialize(), function (result) {
				//and after the ajax request ends we check the text returned
				if (result == 'sent') {
					//if the mail is sent remove the submit paragraph
					$('#cf-submit').remove();
					//and show the mail success div with fadeIn
					$('#mail-success').fadeIn(500);
				} else {
					//show the mail failed div
					$('#mail-fail').fadeIn(500);
					//re enable the submit button by removing attribute disabled and change the text back to Send The Message
					$('#contact-submit').removeAttr('disabled').attr('value', 'Send The Message');
				}
			});
		}
	});

});
// End Jquery Function


	/* ========================================================================= */
	/*	Animated section
	/* ========================================================================= */

	var wow = new WOW(
		{
		  offset:       100,          // distance to the element when triggering the animation (default is 0)
		  mobile:       false      // trigger animations on mobile devices (default is true)
		}
	 );
	 wow.init();


	/* ========================================================================= */
	/*	Smooth Scroll
	/* ========================================================================= */
	var scroll = new SmoothScroll('a[href*="#"]');



	/* ========================================================================= */
	/*	Google Map Customization
	/* =========================================================================  */

	// function initialize() {

	// 	var myLatLng = new google.maps.LatLng(44.738574, -93.1965338);

	// 	var roadAtlasStyles = 

	// 			[
	// 		{
	// 			"featureType": "all",
	// 			"elementType": "labels.text.fill",
	// 			"stylers": [
	// 				{
	// 					"saturation": 100
	// 				},
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 30
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "all",
	// 			"elementType": "labels.text.stroke",
	// 			"stylers": [
	// 				{
	// 					"visibility": "on"
	// 				},
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 6
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "all",
	// 			"elementType": "labels.icon",
	// 			"stylers": [
	// 				{
	// 					"visibility": "on"
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "administrative",
	// 			"elementType": "geometry.fill",
	// 			"stylers": [
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 10
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "administrative",
	// 			"elementType": "geometry.stroke",
	// 			"stylers": [
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 7
	// 				},
	// 				{
	// 					"weight": 1.2
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "landscape",
	// 			"elementType": "geometry",
	// 			"stylers": [
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 0
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "poi",
	// 			"elementType": "geometry",
	// 			"stylers": [
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 3
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "road.highway",
	// 			"elementType": "geometry.fill",
	// 			"stylers": [
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 3
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "road.highway",
	// 			"elementType": "geometry.stroke",
	// 			"stylers": [
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 5
	// 				},
	// 				{
	// 					"weight": 0.2
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "road.arterial",
	// 			"elementType": "geometry",
	// 			"stylers": [
	// 				{
	// 					"color": "#000c17"
	// 				},
	// 				{
	// 					"lightness": 3
	// 				}
	// 			]
	// 		},
	// 		{
				
	// 			"featureType": "road.local",
	// 			"elementType": "geometry",
	// 			"stylers": [
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 6
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "transit",
	// 			"elementType": "geometry",
	// 			"stylers": [
	// 				{
	// 					"color": "#001930"
	// 				},
	// 				{
	// 					"lightness": 7
	// 				}
	// 			]
	// 		},
	// 		{
	// 			"featureType": "water",
	// 			"elementType": "geometry",
	// 			"stylers": [
	// 				{
	// 					"color": "#57cbcc"
	// 				},
	// 				{
	// 					"lightness": 0
	// 				}
	// 			]
	// 		}
	// 	];

	// 	var mapOptions = {
	// 		zoom: 14,
	// 		center: myLatLng,
	// 		disableDefaultUI: true,
	// 		scrollwheel: false,
	// 		navigationControl: false,
	// 		mapTypeControl: false,
	// 		scaleControl: false,
	// 		draggable: false,
	// 		mapTypeControlOptions: {
	// 			mapTypeIds: [google.maps.MapTypeId.ROADMAP, 'roadatlas']
	// 		}
	// 	};

	// 	var map = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);


	// 	var marker = new google.maps.Marker({
	// 		position: myLatLng,
	// 		map: map,
	// 		title: '',
	// 	});


	// 	google.maps.event.addListener(marker, 'click', function () {
	// 		infowindow.open(map, marker);
	// 	});

	// 	var styledMapOptions = {
	// 		name: 'US Road Atlas'
	// 	};

	// 	var usRoadMapType = new google.maps.StyledMapType(
	// 		roadAtlasStyles, styledMapOptions);

	// 	map.mapTypes.set('roadatlas', usRoadMapType);
	// 	map.setMapTypeId('roadatlas');
	// }

	// google.maps.event.addDomListener(window, "load", initialize);

	// ===========================================================================
			// Hero Area Responsiveness:
	// ===========================================================================

	var windowSize = $(window).width();
	if (windowSize < 1200){
		document.getElementById("hero-area-section").classList.remove('hero-area');
		// document.getElementById("hero-area-section").setAttribute("style", "display: none;");
		// document.getElementById("test").setAttribute("style", "display: none;");




	}
	// =====================================================================
	//                 Slideshow:
	// =====================================================================

// 	$("#slideshow > div:gt(0)").hide();

// setInterval(function() { 
//   $('#slideshow > div:first')
//     .fadeOut(1000)
//     .next()
//     .fadeIn(1000)
//     .end()
//     .appendTo('#slideshow');
// },  10000);

	// =====================================================================
	//                Accordian:
	// =====================================================================

	var coll = document.getElementsByClassName("collapsible");
	var i;
	
	for (i = 0; i < coll.length; i++) {
	  coll[i].addEventListener("click", function() {
		this.classList.toggle("active");
		var collapaibleContent = this.nextElementSibling;
		if (collapaibleContent.style.maxHeight){
			collapaibleContent.style.maxHeight = null;
		} else {
			collapaibleContent.style.maxHeight = collapaibleContent.scrollHeight + "px";
		} 
	  });
	}