const imgSelect = document.querySelector('.custom-file-upload');

const profileForm = document.querySelector('#profile-settings-form');
const profileInput = document.querySelector('#upload-excel');

const fileName = document.querySelector('#file-upload-name');
const fileType = document.querySelector('#file-upload-type');

const profileButton = document.querySelector('#profile-submit');
const profileImage = document.querySelector('#profile-settings-image');
const imageUrl = document.querySelector('#profile-image-url');

const statusElement = document.querySelector('#profile-upload-status');
const initialsElement = document.querySelector('.user-profile-initials.settings-profile');

imgSelect.addEventListener('mousedown', () => {
    imgSelect.style.backgroundColor = '#0a2369';
});

imgSelect.addEventListener('mouseup', () => {
    imgSelect.style.backgroundColor = '#0a236900';
});

profileButton.addEventListener('mousedown', () => {
    imgSelect.style.backgroundColor = '#0a2369';
});

profileButton.addEventListener('mouseup', () => {
    imgSelect.style.backgroundColor = '#0a236900';
});

profileForm.onsubmit = (e) => {
    e.preventDefault();
}

profileInput.onchange = (event) => {
    statusElement.innerHTML = '';
    statusElement.opacity = 0;

    fileName.innerHTML = profileInput.files[0].name;
    fileType.innerHTML = profileInput.files[0].type;

    fileName.style.opacity = 1;
    fileType.style.opacity = 1;

    setTimeout(() => {
        // this.settings.preview_image(profileInput, profileImage, imageUrl, initialsElement);
        profileButton.style.display = 'inline-block';
        profileButton.style.opacity = 1;
    }, 400);
}

document.querySelector('#profile-submit').addEventListener('click', (e) => {
    let formData = new FormData(profileForm);
    console.log([...formData])
    profileImageHandler(formData);
    profileInput.value = '';
    fileName.style.opacity = 0;
    fileType.style.opacity = 0;

    profileButton.style.opacity = 0;
    setTimeout(() => {
        fileName.innerHTML = '';
        fileType.innerHTML = '';

        profileButton.style.display = 'none';
    }, 400);
})

function profileImageHandler(formData) {
    // formData.append("unique_id", this.UserManager.user_id)

    // this.UserManager.updateProfileImage(formData)
    //     .then((data) => {
    //         const statusElement = document.querySelector('#profile-upload-status');
    //         setTimeout(() => {
    //             if (data.update_successful) {
    //                 const new_image_url = `${url.location}${url.root}LiaisonWidget/Model/InternalChat/php/images/${data.image_name}`;
    //                 statusElement.innerHTML = 'Profile updated!';
    //                 statusElement.style.color = 'rgb(49, 253, 60)';
    //                 statusElement.style.opacity = 1;

    //                 const profileImage = document.querySelector('#profile-settings-image');
    //                 const imageUrl = document.querySelector('#profile-image-url');
    //                 imageUrl.value = new_image_url;
    //                 profileImage.src = new_image_url;
    //                 this.UserManager.img_url = new_image_url;
    //                 socket.updateUserProfile();
    //             } else {
    //                 statusElement.innerHTML = `${data.message}`;
    //                 statusElement.style.color = 'rgb(253, 49, 49)';
    //                 statusElement.style.opacity = 1;
    //             }
    //         }, 400);
    //     })
}