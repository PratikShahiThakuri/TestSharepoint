import styles from './JsEmployeeSpotlightWebPart.module.scss'
export class SliderHelper{
    public static sliderIndex:number = 0;
    public constructor (){
        SliderHelper.sliderIndex=0;
    }

    /**Auto play slider */
    public startAutoPlay():void{
        var slides = <HTMLScriptElement[]><any>document.getElementsByClassName(styles.mySlides);
        debugger;
        if(slides.length>0){
            for(var i:number=0; i<slides.length;i++){
                slides[i].style.display= SliderHelper.sliderIndex==i?"block":"none";

            }
            SliderHelper.sliderIndex=(++SliderHelper.sliderIndex)>=slides.length?0:SliderHelper.sliderIndex;

        };
    }
        /**move slide */
        public moveSlides(n:number=0):boolean{
            SliderHelper.sliderIndex+=n;
            var slides = (<HTMLScriptElement[]><any>document.getElementsByClassName(styles.mySlides));
            if(slides.length>0){
                if(SliderHelper.sliderIndex>=slides.length){
                    SliderHelper.sliderIndex=0;
                };
                if(SliderHelper.sliderIndex<0){SliderHelper.sliderIndex=slides.length-1;}
                for(var i:number=0; i<slides.length;i++){
                    slides[i].style.display=SliderHelper.sliderIndex==i?"block":"none";
                }
                return true;
            };
        }

    }
