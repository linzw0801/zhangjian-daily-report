// 探测可用的 libsodium 加密库
try { require('tweetnacl'); console.log('tweetnacl: yes'); } catch (e) { console.log('tweetnacl: no'); }
try { require('libsodium-wrappers'); console.log('libsodium-wrappers: yes'); } catch (e) { console.log('libsodium-wrappers: no'); }
try { require('@actions/core'); console.log('@actions/core: yes'); } catch (e) { console.log('@actions/core: no'); }
