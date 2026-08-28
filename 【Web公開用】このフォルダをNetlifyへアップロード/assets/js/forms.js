document.querySelectorAll('[data-application-form]').forEach(form=>{const status=form.querySelector('.form-status');form.addEventListener('submit',event=>{const endpoint=form.dataset.endpoint||'';if(!endpoint||endpoint.includes('YOUR_GOOGLE_APPS_SCRIPT_URL')){event.preventDefault();status.className='form-status error show';status.textContent='現在は受付準備中です。Google Apps ScriptのURL設定後に送信できるようになります。';status.scrollIntoView({behavior:'smooth',block:'center'});return}status.className='form-status success show';status.textContent='送信しています。画面を閉じずにお待ちください。';setTimeout(()=>{status.innerHTML='<strong>送信が完了しました！ありがとうございます！</strong><br>ご登録メールアドレスに申込完了メールが届きますのでご確認ください。届かない場合は、メールアドレスが間違っている可能性もありますので、一度事務局までご確認をお願いいたします。';form.reset()},1800)})});

document.querySelectorAll('input[name="postalCode"]').forEach((postalInput,index)=>{
  const form=postalInput.form;
  const addressInput=form&&form.elements.address;
  if(!addressInput)return;
  const message=document.createElement('small');
  message.className='postal-status';
  message.textContent='7桁入力すると住所を自動検索します。';
  postalInput.insertAdjacentElement('afterend',message);
  let timer;
  const search=()=>{
    const zipcode=postalInput.value.replace(/\D/g,'');
    if(zipcode.length!==7){postalInput.dataset.lastSearched='';message.textContent='7桁入力すると住所を自動検索します。';message.className='postal-status';return}
    if(postalInput.dataset.lastSearched===zipcode)return;
    postalInput.dataset.lastSearched=zipcode;
    postalInput.value=zipcode.slice(0,3)+'-'+zipcode.slice(3);
    message.textContent='住所を検索しています…';message.className='postal-status searching';
    const callback='maunaleiZipCallback_'+index+'_'+Date.now();
    const script=document.createElement('script');
    const cleanup=()=>{delete window[callback];script.remove()};
    window[callback]=data=>{
      if(data&&data.results&&data.results.length){const result=data.results[0];addressInput.value=result.address1+result.address2+result.address3;message.textContent='住所を自動入力しました。番地・建物名を続けて入力してください。';message.className='postal-status success';addressInput.focus();addressInput.setSelectionRange(addressInput.value.length,addressInput.value.length)}else{message.textContent='住所が見つかりませんでした。住所欄へ直接入力してください。';message.className='postal-status error'}
      cleanup();
    };
    script.onerror=()=>{message.textContent='住所を検索できませんでした。住所欄へ直接入力してください。';message.className='postal-status error';cleanup()};
    script.src='https://zipcloud.ibsnet.co.jp/api/search?zipcode='+zipcode+'&callback='+encodeURIComponent(callback);
    document.head.appendChild(script);
  };
  postalInput.addEventListener('input',()=>{clearTimeout(timer);timer=setTimeout(search,500)});
  postalInput.addEventListener('blur',search);
});
