from django.core.exceptions import PermissionDenied

class ManagerRequiredMixin:
    def dispatch(self, request, *args, **kwargs):
        if not request.user.is_manager:
            raise PermissionDenied("دسترسی فقط برای مسئول درمانگاه مجاز است.")
        return super().dispatch(request, *args, **kwargs)

class UserIsOwnerMixin:
    def dispatch(self, request, *args, **kwargs):
        obj = self.get_object()
        if obj.group != request.user.group:
            raise PermissionDenied("شما اجازه دسترسی به این محتوا را ندارید.")
        return super().dispatch(request, *args, **kwargs)